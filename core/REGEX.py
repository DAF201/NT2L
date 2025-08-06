import re
from collections.abc import Iterator


def re_find_first(pattern: str, source: str):
    """find the first apperance of a pattern in string"""
    res = re.search(pattern, source)
    if res:
        return res.group(0)
    else:
        return None


def re_find_all(pattern: str, source: str) -> Iterator:
    """find a pattern in a string, return an iterator"""
    res = re.finditer(pattern, source)
    if res:
        return res
    else:
        return None


def re_compare(pattern: str, source: str) -> bool:
    """compare to check if the given string is in wanted format"""
    res = re.match(pattern, source)
    if res:
        return res.group(0) == source
    return False

