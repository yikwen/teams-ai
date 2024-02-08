"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from abc import ABC, abstractmethod
from typing import Generic, TypeVar, Union

from botbuilder.core import TurnContext

from teams.ai.models.prompt_response import PromptResponse
from teams.ai.planner import Plan
from teams.ai.prompts.sections.prompt_section import PromptSection
from teams.ai.validators.prompt_response_validator import PromptResponseValidator
from teams.state.memory import Memory

ValueT = TypeVar("ValueT")
"Type of message content returned for a 'success' response."


class Augmentation(PromptResponseValidator[ValueT], ABC, Generic[ValueT]):
    """
    An augmentation is a component that can be added to a prompt template to add additional
    functionality to the prompt.
    """

    @abstractmethod
    def create_prompt_section(self) -> Union[PromptSection, None]:
        """
        Creates an optional prompt section for the augmentation.

        Returns:
        Union[PromptSection, None]: The prompt section.
        """

    @abstractmethod
    async def create_plan_from_response(
        self, turn_context: TurnContext, memory: Memory, response: PromptResponse[ValueT]
    ) -> Plan:
        """
        Creates a plan given validated response value.

        Args:
            turn_context (TurnContext): Context for the current turn of conversation.
            memory (Memory): An interface for accessing state variables.
            response (PromptResponse[ValueT]): Validated, transformed response for the prompt.

        Returns:
            Plan: The created plan.
        """
