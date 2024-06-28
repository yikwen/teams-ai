"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from logging import Logger
from typing import Any, Awaitable, Callable, Dict, List, Optional, Union, cast

import openai
from botbuilder.core import TurnContext
from openai.types import chat, shared_params

from ...state import MemoryBase
from ..prompts.message import Message, MessageContext
from ..prompts.prompt_functions import PromptFunctions
from ..prompts.prompt_template import PromptTemplate
from ..tokenizers import Tokenizer
from .openai_function import OpenAIFunction
from .prompt_completion_model import PromptCompletionModel
from .prompt_response import PromptResponse


@dataclass
class OpenAIModelOptions:
    """
    Options for configuring an `OpenAIModel` to call an OpenAI hosted model.
    """

    api_key: str
    "API key to use when calling the OpenAI API."

    default_model: str
    "Default model to use for completions."

    endpoint: Optional[str] = None
    "Optional. Endpoint to use when calling the OpenAI API."

    organization: Optional[str] = None
    "Optional. Organization to use when calling the OpenAI API."

    logger: Optional[Logger] = None
    "Optional. When set the model will log requests"


@dataclass
class AzureOpenAIModelOptions:
    """
    Options for configuring an `OpenAIModel` to call an Azure OpenAI hosted model.
    """

    default_model: str
    "Default name of the Azure OpenAI deployment (model) to use."

    endpoint: str
    "Deployment endpoint to use."

    api_version: str = "2023-05-15"
    "Optional. Version of the API being called. Defaults to `2023-05-15`."

    api_key: Optional[str] = None
    "API key to use when making requests to Azure OpenAI."

    azure_ad_token_provider: Optional[Callable[..., str]] = None
    """Optional. A function that returns an access token for Microsoft Entra 
    (formerly known as Azure Active Directory), which will be invoked in every request.
    """

    organization: Optional[str] = None
    "Optional. Organization to use when calling the OpenAI API."

    logger: Optional[Logger] = None
    "Optional. When set the model will log requests"


class OpenAIModel(PromptCompletionModel):
    """
    A `PromptCompletionModel` for calling OpenAI and Azure OpenAI hosted models.
    """

    _options: Union[OpenAIModelOptions, AzureOpenAIModelOptions]
    _client: openai.AsyncOpenAI

    @property
    def options(self) -> Union[OpenAIModelOptions, AzureOpenAIModelOptions]:
        return self._options

    def __init__(self, options: Union[OpenAIModelOptions, AzureOpenAIModelOptions]) -> None:
        """
        Creates a new `OpenAIModel` instance.

        Args:
            options (OpenAIModelOptions | AzureOpenAIModelOptions): model options.
        """

        self._options = options

        if isinstance(options, OpenAIModelOptions):
            self._client = openai.AsyncOpenAI(
                api_key=options.api_key,
                base_url=options.endpoint,
                organization=options.organization,
                default_headers={"User-Agent": self.user_agent},
            )
        elif isinstance(options, AzureOpenAIModelOptions):
            self._client = openai.AsyncAzureOpenAI(
                api_key=options.api_key,
                api_version=options.api_version,
                azure_ad_token_provider=options.azure_ad_token_provider,
                azure_endpoint=options.endpoint,
                azure_deployment=options.default_model,
                organization=options.organization,
                default_headers={"User-Agent": self.user_agent},
            )

    async def complete_prompt(
        self,
        context: TurnContext,
        memory: MemoryBase,
        functions: PromptFunctions,
        tokenizer: Tokenizer,
        template: PromptTemplate,
    ) -> PromptResponse[str]:
        # pylint: disable=too-many-locals
        max_tokens = template.config.completion.max_input_tokens

        # Setup tools if enabled
        is_tools_enabled = template.config.completion.include_tools
        tool_choice = template.config.completion.tool_choice
        parallel_tool_calls = template.config.completion.parallel_tool_calls
        tools: List[OpenAIFunction] = []
        tools_handlers = memory.get("temp.tools")

        # If tools is enabled, reformat actions to appropriate schema
        if is_tools_enabled:
            if not template.actions:
                return PromptResponse[str](status="tools_error", error="Missing actions")
            if not tools_handlers or len(tools_handlers) == 0:
                return PromptResponse[str](status="tools_error", error="Missing tools handlers")
            if len(template.actions) != len(tools_handlers):
                return PromptResponse[str](
                    status="tools_error",
                    error="Number of actions does not match number of tool handlers",
                )
            for action in template.actions:
                if action.name in tools_handlers:
                    handler = tools_handlers.get(action.name).func
                    tool = OpenAIFunction(
                        action.name, action.description, action.parameters, handler
                    )
                    tools.append(tool)

        formatted_tools: List[chat.ChatCompletionToolParam] = []

        for tool in tools:
            curr_tool = chat.ChatCompletionToolParam(
                type="function",
                function=shared_params.FunctionDefinition(
                    name=tool.name,
                    description=tool.description or "",
                    parameters=tool.parameters or {},
                ),
            )
            formatted_tools.append(curr_tool)

        model = (
            template.config.completion.model
            if template.config.completion.model is not None
            else self._options.default_model
        )

        res = await template.prompt.render_as_messages(
            context=context,
            memory=memory,
            functions=functions,
            tokenizer=tokenizer,
            max_tokens=max_tokens,
        )

        if res.too_long:
            return PromptResponse[str](
                status="too_long",
                error=f"""
                the generated chat completion prompt had a length of {res.length} tokens
                which exceeded the max_input_tokens of {max_tokens}
                """,
            )

        if self._options.logger is not None:
            self._options.logger.debug(f"PROMPT:\n{res.output}")

        messages: List[chat.ChatCompletionMessageParam] = []

        for msg in res.output:
            param: Union[
                chat.ChatCompletionUserMessageParam,
                chat.ChatCompletionAssistantMessageParam,
                chat.ChatCompletionSystemMessageParam,
            ] = chat.ChatCompletionUserMessageParam(
                role="user",
                content=msg.content if msg.content is not None else "",
            )

            if msg.name:
                param["name"] = msg.name

            if msg.role == "assistant":
                param = chat.ChatCompletionAssistantMessageParam(
                    role="assistant",
                    content=msg.content if msg.content is not None else "",
                )

                if msg.name:
                    param["name"] = msg.name
            elif msg.role == "system":
                param = chat.ChatCompletionSystemMessageParam(
                    role="system",
                    content=msg.content if msg.content is not None else "",
                )

                if msg.name:
                    param["name"] = msg.name

            messages.append(param)

        try:
            extra_body = {}
            if template.config.completion.data_sources is not None:
                extra_body["data_sources"] = template.config.completion.data_sources
            completion = await self._client.chat.completions.create(
                messages=messages,
                model=model,
                presence_penalty=template.config.completion.presence_penalty,
                frequency_penalty=template.config.completion.frequency_penalty,
                top_p=template.config.completion.top_p,
                temperature=template.config.completion.temperature,
                max_tokens=max_tokens,
                tools=formatted_tools,
                tool_choice=tool_choice or "auto",
                parallel_tool_calls=parallel_tool_calls or True,
                extra_body=extra_body,
            )

            if self._options.logger is not None:
                self._options.logger.debug("COMPLETION:\n%s", completion.model_dump_json())

            # Handle tools flow
            response_message = completion.choices[0].message
            tool_calls = response_message.tool_calls

            # Tracks the latest response from the LLM
            final_response = completion
            augmentation = template.config.augmentation

            if (
                augmentation
                and augmentation.augmentation_type == "none"
                and is_tools_enabled
                and tool_calls
            ):
                while tool_calls and len(tool_calls) > 0:
                    if not parallel_tool_calls and len(tool_calls) > 1:
                        break

                    if isinstance(tool_choice, dict) and len(tool_calls) > 1:
                        break

                    if tool_choice == "none":
                        break

                    messages.append(
                        cast(chat.ChatCompletionAssistantMessageParam, response_message)
                    )

                    if isinstance(tool_choice, dict):
                        # Calling a single tool
                        function_name = tool_choice["function"]["name"]
                        curr_tool_call = tool_calls[0]
                        curr_function = next(tool for tool in tools if tool.name == function_name)

                        # Validate function name
                        if not curr_function:
                            break

                        # Validate function arguments
                        required_args = (
                            curr_function.parameters["required"]
                            if curr_function.parameters and "required" in curr_function.parameters
                            else None
                        )

                        curr_args = json.loads(curr_tool_call.function.arguments)
                        curr_function_handler = curr_function.handler

                        if required_args:
                            if len(required_args) > len(curr_args):
                                break

                        # Call the function
                        function_response = await self._handle_function_response(
                            curr_function_handler, curr_args
                        )

                        messages.append(
                            chat.ChatCompletionToolMessageParam(
                                role="tool",
                                tool_call_id=curr_tool_call.id,
                                content=function_response,
                            )
                        )
                    else:
                        curr_message_length = len(messages)
                        messages = await self._handle_multiple_tool_calls(
                            messages, tool_calls, tools
                        )
                        # No tools were run successfully
                        if len(messages) == curr_message_length:
                            break

                    final_response = await self._client.chat.completions.create(
                        messages=messages,
                        model=model,
                        presence_penalty=template.config.completion.presence_penalty,
                        frequency_penalty=template.config.completion.frequency_penalty,
                        top_p=template.config.completion.top_p,
                        temperature=template.config.completion.temperature,
                        max_tokens=max_tokens,
                    )

                    tool_calls = final_response.choices[0].message.tool_calls

            input: Optional[Message] = None
            last_message = len(res.output) - 1

            # Skips the first message which is the prompt
            if last_message > 0 and res.output[last_message].role == "user":
                input = res.output[last_message]

            return PromptResponse[str](
                input=input,
                message=Message(
                    role=final_response.choices[0].message.role,
                    content=final_response.choices[0].message.content,
                    context=(
                        MessageContext.from_dict(final_response.choices[0].message.context)
                        if hasattr(final_response.choices[0].message, "context")
                        else None
                    ),
                ),
            )
        except openai.APIError as err:
            if self._options.logger is not None:
                self._options.logger.error("ERROR:\n%s", json.dumps(err.body))

            return PromptResponse[str](
                status="error",
                error=f"""
                The chat completion API returned an error
                status of {err.code}: {err.message}
                """,
            )

    async def _handle_function_response(
        self,
        curr_function_handler: Callable[..., Awaitable[str]],
        curr_args: Dict[str, Any],
    ) -> str:
        if len(curr_args) > 0:
            return await curr_function_handler(**curr_args)
        return await curr_function_handler()

    async def _handle_multiple_tool_calls(
        self,
        messages: List[chat.ChatCompletionMessageParam],
        tool_calls: List[chat.ChatCompletionMessageToolCall],
        tools: List[OpenAIFunction],
    ) -> List[chat.ChatCompletionMessageParam]:
        for tool_call in tool_calls:
            tool_name = tool_call.function.name
            curr_function = next(tool for tool in tools if tool.name == tool_name)

            # Validate function name
            if not curr_function:
                continue

            # Validate function arguments
            required_args = (
                curr_function.parameters["required"]
                if curr_function.parameters and "required" in curr_function.parameters
                else None
            )

            curr_args = json.loads(tool_call.function.arguments)
            curr_function_handler = curr_function.handler

            if required_args:
                if len(required_args) > len(curr_args):
                    continue

            # Call the function
            function_response = await self._handle_function_response(
                curr_function_handler, curr_args
            )

            messages.append(
                chat.ChatCompletionToolMessageParam(
                    role="tool",
                    tool_call_id=tool_call.id,
                    content=function_response,
                )
            )

        return messages
