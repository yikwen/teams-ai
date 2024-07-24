"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Generic, List, Literal, Optional, TypeVar, Union

from dataclasses_json import DataClassJsonMixin, dataclass_json

from .function_call import FunctionCall

T = TypeVar("T")


@dataclass_json
@dataclass
class Message(Generic[T], DataClassJsonMixin):
    """
    A message object sent to or received from an LLM.

    Attributes:
        role (str): The messages role. Typically 'system', 'user', 'assistant', 'function'.
        content (Optional[T]): Text of the message.
        function_call (Optional[FunctionCall]): A named function to call.
        name (Optional[str]): Name of the function that was called.
        action_calls (Optional[List[ActionCall]]):
            List of tool calls from the model.
    """

    role: str
    content: Optional[T] = None
    context: Optional[MessageContext] = None
    function_call: Optional[FunctionCall] = None
    name: Optional[str] = None
    action_calls: Optional[List[ActionCall]] = None


@dataclass
class ImageUrl:
    """
    Url for an image

    Attributes:
        url (str): Url of the image
    """

    url: str


@dataclass
class TextContentPart:
    """
    Represents text content part of a message

    Attributes:
        text (str): Text content
    """

    type: Literal["text"]
    text: str


@dataclass
class ImageContentPart:
    """
    Represents image content part of a message

    Attributes:
        image_url (Union[str, ImageUrl]): Url for the image
    """

    type: Literal["image_url"]
    image_url: Union[str, ImageUrl]


MessageContentParts = Union[TextContentPart, ImageContentPart]
"""
Represents part of the message's content
"""


@dataclass_json
@dataclass
class Citation(DataClassJsonMixin):
    """
    Represents a citation returned by the model

    Attributes:
        content (str): The content of the citation
        title (str): The title of the citation
        url (str): The url of the citation
        filepath (str): The filepath of the citation
    """

    content: str
    title: Optional[str]
    url: Optional[str]
    filepath: Optional[str]


@dataclass_json
@dataclass
class MessageContext(DataClassJsonMixin):
    """
    Represents the message context containing a citation

    Attributes:
        citations (list[Citation]): The citations in the message
        intent (str): The intent of the message
    """

    citations: list[Citation]
    intent: str


@dataclass_json
@dataclass
class ActionCall(DataClassJsonMixin):
    """
    Action called by the model.
    """

    id: str
    """The ID of the action call."""

    function: ActionFunction
    """The function that the model called."""

    type: Literal["function"]
    """The type of the action. Currently, only `function` is supported."""


@dataclass_json
@dataclass
class ActionFunction(DataClassJsonMixin):
    """
    Function details associated with an action called by a model.
    """

    arguments: str
    """
    The arguments to call the function with, as generated by the model in JSON
    format.
    """

    name: str
    """The name of the action to call."""
