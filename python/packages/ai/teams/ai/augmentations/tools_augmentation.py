"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from __future__ import annotations

import json
from typing import List, Optional, Union, cast

from botbuilder.core import TurnContext

from ...state import MemoryBase
from ..models.chat_completion_action import ChatCompletionAction
from ..models.prompt_response import PromptResponse
from ..planners.plan import (
    Plan,
    PredictedCommand,
    PredictedDoCommand,
    PredictedSayCommand,
)
from ..prompts.message import ActionCall, Message
from ..prompts.sections.prompt_section import PromptSection
from ..tokenizers.tokenizer import Tokenizer
from ..validators.validation import Validation
from .augmentation import Augmentation
from .tools_constants import ACTIONS_HISTORY


class ToolsAugmentation(Augmentation[Union[str, List[ActionCall]]]):
    """
    A server-side 'tools' augmentation.
    """

    _actions: Optional[List[ChatCompletionAction]]

    def __init__(self, actions: Optional[List[ChatCompletionAction]]) -> None:
        self._actions = actions

    def create_prompt_section(self) -> Optional[PromptSection]:
        """
        Creates an optional prompt section for the augmentation.
        """
        return None

    async def validate_response(
        self,
        context: TurnContext,
        memory: MemoryBase,
        tokenizer: Tokenizer,
        response: PromptResponse[Union[str, List[ActionCall]]],
        remaining_attempts: int,
    ) -> Validation:
        """
        Validates a response to a prompt.

        Args:
            context (TurnContext): Context for the current turn of conversation.
            memory (MemoryBase): Interface for accessing state variables.
            tokenizer (Tokenizer): Tokenizer to use for encoding/decoding text.
            response (PromptResponse[Union[str, List[ActionCall]]]):
                Response to validate.
            remaining_attempts (int): Nubmer of remaining attempts to validate the response.

        Returns:
            Validation: A 'Validation' object.
        """
        if response.message and response.message.action_calls and memory.has(ACTIONS_HISTORY):
            tool_calls = response.message.action_calls
            tools = self._actions

            if tools and len(tool_calls) > 0:
                return Validation(
                    valid=True,
                    value=tool_calls,
                )

        memory.set(ACTIONS_HISTORY, [])
        return Validation(valid=True)

    async def create_plan_from_response(
        self,
        turn_context: TurnContext,
        memory: MemoryBase,
        response: PromptResponse[Union[str, List[ActionCall]]],
    ) -> Plan:
        """
        Creates a plan given validated response value.

        Args:
            turn_context (TurnContext): Context for the current turn of conversation.
            memory (MemoryBase): Interface for accessing state variables.
            response (PromptResponse[Union[str, List[ActionCall]]]):
                The validated and transformed response for the prompt.

        Returns:
            Plan: The created plan.
        """

        commands: List[PredictedCommand] = []

        if response.message and response.message.content:
            if memory.has(ACTIONS_HISTORY) and isinstance(response.message.content, list):
                tool_calls: List[ActionCall] = response.message.content

                for tool in tool_calls:
                    command = PredictedDoCommand(
                        action=tool.function.name,
                        parameters=json.loads(tool.function.arguments),
                    )
                    commands.append(command)
                return Plan(commands=commands)
            say_response = cast(Message[str], response.message)
            return Plan(commands=[PredictedSayCommand(response=say_response)])
        return Plan()
