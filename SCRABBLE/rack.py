"""
This module contains the functions regarding the rack.
Functions will be used in this module itself and be imported
into scrabble.py.
"""

import re
from string import ascii_lowercase
from itertools import combinations_with_replacement

score_dict = {"a": 1, "c": 3, "b": 3, "e": 1, "d": 2, "g": 2,
            "f": 4, "i": 1, "h": 4, "k": 5, "j": 8, "m": 3,
            "l": 1, "o": 1, "n": 1, "q": 10, "p": 3, "s": 1,
            "r": 1, "u": 1, "t": 1, "w": 4, "v": 4, "y": 4,
            "x": 8, "z": 10, '*':0, '?':0}

def check_rack(rack:str):
    """Check the rack string input

    Args:
        rack (str): one single string of the tiles, such as 'PEN*?in'

    Returns:
        _type_: _description_
    """
    chr_regex = re.compile('[^A-Za-z*?]')
    if len(rack) == 0:
        message = 'No input detected'
    elif len(rack) == 1 or len(rack) > 7:
        message = 'Enter a rack with length of 2-7'
    elif re.search(chr_regex, rack):
        message = 'Invalid string input'
    elif rack.count('*') > 1 or rack.count('?') > 1:
        message = 'Only one of each wildcard is allowed'
    else:
        message = None
    return message

def fill_wildcard(rack:str) -> list:
    """
    Re-configure the wildcard(s) with alphabet(s). 
    Return a list containing all possible combinations of the rack.

    Args:
        rack (str): one single string of the tiles, such as 'PEN*?in'

    Returns:
        list: A list of racks that already accounted for all letter combinations
        such as [['a','b','b'],['a','b','c']]
    """
    # Sorting the rack is important so the wildcards are in certain positions
    sorted_rack = sorted(list(rack.lower()))
    alphabet = list(ascii_lowercase)

    if (sorted_rack[0] == '*' and not "?" in sorted_rack) or\
        (sorted_rack[0] == '?' and not "*" in sorted_rack):
        rack_combs = [sorted([letter] + sorted_rack[1:]) for letter in alphabet]
    elif sorted_rack[0] == '*' and sorted_rack[1] == '?':
        combs = list(combinations_with_replacement(alphabet, 2))
        rack_combs = [sorted(list(comb) + sorted_rack[2:]) for comb in combs]
    else:
        rack_combs = [sorted_rack]
    return rack_combs

def find_best_score(word: str, rack:str):
    """
        Given a word and original rack, find the best way to form
        the word and calculate that highest score. (You check if you
        have actual letters first before using wildcard.) 
    Args:
        word (str): a word in string 
        rack (str): the tiles player get without accounting for wildcards 
    Returns:
        _type_: _description_
    """
    original = list(rack.lower())
    used_characters = []
    for letter in word:
        try:
            index_original = original.index(letter)
            original.pop(index_original)
            used_characters.append(letter)
        except ValueError:
            if '*' in original:
                original.pop(original.index('*'))
                used_characters.append('*')
            elif '?' in original:
                original.pop(original.index('?'))
                used_characters.append('?')
        except Exception as ex:
            raise ex
        finally:
            score = sum([score_dict[character] for character in used_characters])
    return score, used_characters
