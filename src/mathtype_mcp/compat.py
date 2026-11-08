"""Catch malformed braces before invoking the native parser; not a TeX compiler."""


def check_input(source: str) -> None:
    depth = 0
    index = 0
    while index < len(source):
        char = source[index]
        if char == "\\":
            index += 2
            continue
        if char == "%":
            raise ValueError("TeX comments are not supported; escape a literal percent sign.")
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth < 0:
                raise ValueError("Unmatched closing brace.")
        index += 1
    if depth:
        raise ValueError("Unmatched opening brace.")
