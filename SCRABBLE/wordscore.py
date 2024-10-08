"""
This module contains the functions regarding finding match from
the list and calculate the score.
Functions will be used in this module itself and be imported
into scrabble.py
"""

import bisect
from rack import find_best_score

def get_char_freq(word) -> dict:
    '''
    Return a dictionary that records the frequency of each character in a word
    '''
    char_freq_dict = {}
    for letter in word:
        char_freq_dict[letter] = char_freq_dict.get(letter, 0) + 1
    return char_freq_dict

def match_words_w_wildcard_2(vocabs: list, rack_combs: list) -> list:
    """
    Generate possible vocab list based on the rack. With wildcard
    Args:
        vocabs (list): all words from sowpods after adjusting for
        the length of rack ['aa','ab',..]
        rack_combs (list): all possible combination after filling 
        wildcard [['a','a'], ['b','a'], ['c','a']...]

    Returns:
        list:
    """
    find_index = lambda letter: bisect.bisect_left(vocabs, letter)

    # Precompute character frequencies for vocab
    vocab_freqs = {word: get_char_freq(word) for word in vocabs}

    matches = set()  # Use a set to avoid duplicates
    for rack_comb in rack_combs:
        rack_set = set(rack_comb)  # Convert to set for faster lookups
        start_line_index = find_index(rack_comb[0])

        while start_line_index < len(vocabs):
            word = vocabs[start_line_index]
            char_freq = vocab_freqs[word]
            valid = True

            for char, count in char_freq.items():
                if char not in rack_set or count > rack_comb.count(char):
                    valid = False
                    break  # Break early if invalid

            if valid:
                matches.add(word)  # Add directly to the set

            start_line_index += 1

    return list(matches)

def score_word(matches: list, original_rack:str) -> list:
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
        score, used_characters = find_best_score(word, original_rack)
        scores.append((score, word.upper()))
    scores = sorted(scores, key = lambda x: (-x[0], x[1]))
    scores = [scores, len(scores)]
    return scores
