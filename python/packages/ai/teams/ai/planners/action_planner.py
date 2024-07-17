"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from logging import Logger
from typing import Awaitable, Callable, List, Optional, TypeVar, Union

from botbuilder.core import TurnContext
from openai.types import chat

from ...app_error import ApplicationError
from ...state import MemoryBase, TurnState
from ..augmentations.default_augmentation import DefaultAugmentation
from ..augmentations.tools_constants import (
    SUBMIT_TOOL_OUTPUTS_MAP,
    SUBMIT_TOOL_OUTPUTS_MESSAGES,
    SUBMIT_TOOL_OUTPUTS_VARIABLE,
)
from ..clients import LLMClient, LLMClientOptions
from ..models.prompt_completion_model import PromptCompletionModel
from ..models.prompt_response import PromptResponse
from ..prompts.prompt_functions import PromptFunctions
from ..prompts.prompt_manager import PromptManager
from ..prompts.prompt_template import PromptTemplate
from ..tokenizers import GPTTokenizer, Tokenizer
from ..validators import DefaultResponseValidator, PromptResponseValidator
from .plan import Plan
from .planner import Planner

ActionPlannerPromptFactory = Callable[
    [TurnContext, TurnState, "ActionPlanner"], Awaitable[PromptTemplate]
]

StateT = TypeVar("StateT", bound=TurnState)


@dataclass
class ActionPlannerOptions:
    """
    Options used to configure an `ActionPlanner` instance.
    """

    model: PromptCompletionModel
    "Model instance to use."

    prompts: PromptManager
    "Prompt manager used to manage prompts."

    default_prompt: Union[str, ActionPlannerPromptFactory] = "default"
    "The default prompt to use. Defaults to `default`"

    max_repair_attempts: int = 3
    "Maximum number of repair attempts to make. Defaults to `3`"

    tokenizer: Tokenizer = field(default_factory=GPTTokenizer)
    "Optional tokenizer to use. Defaults to `GPTTokenizer`"

    logger: Optional[Logger] = None
    "Optional. When set the model will log requests"


class ActionPlanner(Planner[StateT]):
    """
    A planner that uses a Large Language Model (LLM) to generate plans.
    """

    _options: ActionPlannerOptions
    _prompt_factory: ActionPlannerPromptFactory

    @property
    def options(self) -> ActionPlannerOptions:
        return self._options

    def __init__(self, options: ActionPlannerOptions) -> None:
        """
        Creates a new `ActionPlanner` instance.

        Args:
            options (ActionPlannerOptions): Options used to configure the planner.
        """

        self._options = options

        if isinstance(self._options.default_prompt, str):
            self._prompt_factory = self._default_prompt_factory(self._options.default_prompt)
        else:
            self._prompt_factory = self._options.default_prompt

    async def begin_task(self, context: TurnContext, state: TurnState) -> Plan:
        return await self.continue_task(context, state)

    async def continue_task(self, context: TurnContext, state: TurnState) -> Plan:
        template = await self._prompt_factory(context, state, self)
        augmentation = template.augmentation or DefaultAugmentation()
        res = await self.complete_prompt(
            context=context, memory=state, prompt=template, validator=augmentation
        )

        if res.status != "success":
            raise ApplicationError(res.error or "[ActionPlanner]: failed task")

        return await augmentation.create_plan_from_response(context, state, res)

    async def complete_prompt(
        self,
        context: TurnContext,
        memory: MemoryBase,
        prompt: Union[str, PromptTemplate],
        validator: PromptResponseValidator = DefaultResponseValidator(),
    ) -> PromptResponse[str]:
        """
        Completes a prompt using an optional validator.

        Args:
            context (TurnContext): The current turn context.
            memory (MemoryBase): A memory interface used to access state variables
                (the turn state object implements this interface.)
            prompt (Union[str, PromptTemplate]): Name of the prompt to use or a prompt template.
            validator (Validator): Optional. A validator to use to validate
                the response returned by the model.
        """
        name = ""

        if isinstance(prompt, str):
            name = prompt
        else:
            name = prompt.name

            if not self._options.prompts.has_prompt(prompt.name):
                self._options.prompts.add_prompt(prompt)

        template = await self._options.prompts.get_prompt(name)
        memory = self._handle_action_tools(memory, template)

        include_history = template.config.completion.include_history
        client = LLMClient(
            LLMClientOptions(
                model=self._options.model,
                history_variable=(
                    f"conversation.{name}_history" if include_history else f"temp.{name}_history"
                ),
                input_variable="temp.input",
                validator=validator,
                logger=self._options.logger,
            )
        )

        return await client.complete_prompt(
            context=context,
            memory=memory,
            functions=self._options.prompts,
            tokenizer=self._options.tokenizer,
            template=template,
        )

    def _handle_action_tools(self, memory: MemoryBase, template: PromptTemplate) -> MemoryBase:
        if not memory.get(SUBMIT_TOOL_OUTPUTS_VARIABLE):
            memory.set("temp.tool_choice", template.config.completion.tool_choice)
            memory.set("temp.parallel_tool_calls", template.config.completion.parallel_tool_calls)

            # Ensure clean state
            memory.set(SUBMIT_TOOL_OUTPUTS_MAP, {})
            memory.set(SUBMIT_TOOL_OUTPUTS_MESSAGES, {})

        elif memory.get(SUBMIT_TOOL_OUTPUTS_VARIABLE) is True:
            # Submit tool outputs
            action_outputs = memory.get("temp.action_outputs") or []
            tool_map = memory.get(SUBMIT_TOOL_OUTPUTS_MAP) or {}
            tool_outputs: List[chat.ChatCompletionToolMessageParam] = []

            for action in action_outputs:
                output = action_outputs[action]
                if tool_map:
                    tool_call_id = tool_map[action] if action in tool_map else None
                    if tool_call_id is not None:
                        tool_outputs.append(
                            chat.ChatCompletionToolMessageParam(
                                tool_call_id=tool_call_id, role="tool", content=output
                            )
                        )

            if len(tool_outputs) > 0:
                memory.set(SUBMIT_TOOL_OUTPUTS_MESSAGES, tool_outputs)
            else:
                # Reset state
                memory.set(SUBMIT_TOOL_OUTPUTS_VARIABLE, False)
                memory.set(SUBMIT_TOOL_OUTPUTS_MAP, {})
                memory.set(SUBMIT_TOOL_OUTPUTS_MESSAGES, {})
        return memory

    def add_semantic_function(
        self, prompt: Union[str, PromptTemplate], _validator: Optional[PromptResponseValidator]
    ) -> "ActionPlanner":
        name = ""

        if isinstance(prompt, PromptTemplate):
            name = prompt.name
            self._options.prompts.add_prompt(prompt)
        else:
            name = prompt

        async def __func__(
            _context: TurnContext,
            memory: MemoryBase,
            _functions: PromptFunctions,
            _tokenizer: Tokenizer,
            args: List[str],
        ):
            memory.set("temp.input", " ".join(args))

        self._options.prompts.add_function(name, __func__)
        return self

    def _default_prompt_factory(self, name: str) -> ActionPlannerPromptFactory:
        async def __factory__(
            _context: TurnContext, _state: TurnState, _planner: "ActionPlanner"
        ) -> PromptTemplate:
            return await self._options.prompts.get_prompt(name)

        return __factory__
