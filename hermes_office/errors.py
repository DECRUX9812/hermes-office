from __future__ import annotations


class OfficeError(Exception):
    reason = "office_error"
    hint = "Check the request and try again."


class PolicyRefusal(OfficeError):
    reason = "policy_refusal"


class OutsideDeskError(PolicyRefusal):
    reason = "outside_desk"
    hint = "Add the folder to the desk first, then retry."


class UnsupportedFileError(PolicyRefusal):
    reason = "unsupported_file"
    hint = "Use a .docx, .pptx, or .xlsx file."


class MissingFolderError(PolicyRefusal):
    reason = "missing_folder"
    hint = "Register the folder before writing to it."


class NotFoundError(OfficeError):
    reason = "not_found"
    hint = "Check the target path or search term."


class AmbiguousSelectionError(OfficeError):
    reason = "ambiguous_target"
    hint = "Narrow the target with nth, story, or a more specific query."
