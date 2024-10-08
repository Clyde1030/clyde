"""
This module contains the functions regarding the rack
Functions will be used in this module itself and be imported
into scrabble.py
"""

import re
from string import ascii_lowercase
import bisect
from itertools import combinations_with_replacement

score_dict = {"a": 1, "c": 3, "b": 3, "e": 1, "d": 2, "g": 2,
            "f": 4, "i": 1, "h": 4, "k": 5, "j": 8, "m": 3,
            "l": 1, "o": 1, "n": 1, "q": 10, "p": 3, "s": 1,
            "r": 1, "u": 1, "t": 1, "w": 4, "v": 4, "y": 4,
            "x": 8, "z": 10, '*':0, '?':0}

def check_rack(rack:str):
    '''Check the rack input'''
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

def sort_rack(rack:str) -> list:
    '''
    Sort the rack alphebatically and return as a list. Make wildcards in the front.
    '''
    sorted_rack = []
    for letter in rack.lower():
        sorted_rack.append(letter)
        sorted_rack.sort()
    return sorted_rack

def get_char_freq(word) -> dict:
    '''
    Return a dictionary that records the frequency of each character in a word
    '''
    char_freq_dict = {}
    for letter in word:
        char_freq_dict[letter] = char_freq_dict.get(letter, 0) + 1
    return char_freq_dict

################################################################
# Here start sections for rack without a wildcard
def match_words(vocabs:list, rack_combs:list) -> list:
    '''
    Generate possible vocab list based on the rack. No wildcard
    '''
    matches = []
    for vocab in vocabs:
        valid = True
        char_freq = get_char_freq(vocab)
        for char in char_freq:
            if char not in rack_combs:
                valid = False
            elif char_freq.get(char) > rack_combs.count(char):
                valid = False
        if valid:
            matches.append(vocab)
            continue
    return matches

def score_word(words:list) -> tuple:
    '''take each word and return the score based on a dictionary'''
    words = sorted(words)
    scores = []
    for word in words:
        score = sum([score_dict[letter] for letter in word])
        scores.append((score, word.upper()))
    scores = sorted(scores, key = lambda x:x[0], reverse=True)
    return (scores, len(scores))
################################################################
# Here start sections for rack with one or two wildcard

def fill_wildcard(rack:list) -> list:
    '''
    Re-configure the wildcard(s) with alphabet(s). 
    Return a list containing all possible combinations of the rack.
    '''
    alphabet = list(ascii_lowercase)
    potential_rack = []
    if (rack[0] == '*' and not "?" in rack) or (rack[0] == '?' and not "*" in rack):
        potential_rack = [sorted([letter] + rack[1:]) for letter in alphabet]
    elif rack[0] == '*' and rack[1] == '?':
        combs = list(combinations_with_replacement(alphabet, 2))
        potential_rack = [sorted(list(comb) + rack[2:]) for comb in combs]
    else:
        potential_rack = sorted(rack)
    return potential_rack

def match_words_w_wildcard(vocabs:list, rack_combs:list) -> list:
    '''
    Generate possible vocab list based on the rack. With wildcard
    vocabs: all words from sowpods after adjusting for the length of rack ['aa','ab',..]
    rack_combs: all possible combination after filling wildcard [['a','a'], ['b','a'], ['c','a']...]
    '''
    matches = []

    for rack_comb in rack_combs:
        for vocab in vocabs:
            valid = True
            char_freq = get_char_freq(vocab)
            for char in char_freq:
                if (char not in rack_comb) or (char_freq[char] > rack_comb.count(char)):
                    valid = False
            if valid:
                matches.append(vocab)
                continue
    matches = list(set(matches))
    return matches # matches_dict

def _find_best_score(word, rack:str):
    """_summary_

    Args:
        word (str): a word in string 
        rack (str): the tiles player get without accounting for wildcards 
    Returns:
        _type_: _description_
    """
    original = [char.lower() for char in rack]
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


def score_words_w_wildcard(matches: list, original_rack:str) -> list:
    """Assign a score to each match in the list

    Args:
        matches (list): a list of available words
        original_rack (str): the tiles player get without accounting for wildcards

    Returns:
        list: a list of tuples (score, word) and the length of elements
        Sorted first by score then by word.
    """
    scores = []
    for word in matches:
        score, used_characters = _find_best_score(word, original_rack)
        scores.append((score, word.upper()))
    scores = sorted(scores, key = lambda x: (-x[0], x[1]))
    scores = [scores, len(scores)]
    return scores

def match_words_w_wildcard_2(vocabs:list, rack_combs:list) -> list:
    find_index = lambda letter: bisect.bisect_left(vocabs, letter)
    matches = []
    for rack_comb in rack_combs:
        start_line_index = find_index(rack_comb[0])
        while start_line_index < len(vocabs):
            valid = True
            char_freq = get_char_freq(vocabs[start_line_index])
            for char in char_freq:
                if (char not in rack_comb) or (char_freq[char] > rack_comb.count(char)):
                    valid = False
            if valid:
                matches.append(vocabs[start_line_index])
            start_line_index += 1
    matches = list(set(matches))
    return matches # matches_dict
    

def match_words_w_wildcard_3(vocabs: list, rack_combs: list) -> list:
    find_index = lambda letter: bisect.bisect_left(vocabs, letter)

    vocab_freqs = {word: get_char_freq(word) for word in vocabs}

    matches = set()
    for rack_comb in rack_combs:
        rack_set = set(rack_comb)
        start_line_index = find_index(rack_comb[0])

        while start_line_index < len(vocabs):
            word = vocabs[start_line_index]
            char_freq = vocab_freqs[word]
            valid = True

            for char, count in char_freq.items():
                if char not in rack_set or count > rack_comb.count(char):
                    valid = False
                    break

            if valid:
                matches.add(word)

            start_line_index += 1
    
    return list(matches)  