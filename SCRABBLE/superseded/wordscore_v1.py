"""
This module contains the functions regarding the rack
Functions will be used in this module itself and be imported
into scrabble.py
"""

import re
from string import ascii_lowercase
from itertools import combinations_with_replacement

score_dict = {"a": 1, "c": 3, "b": 3, "e": 1, "d": 2, "g": 2,
            "f": 4, "i": 1, "h": 4, "k": 5, "j": 8, "m": 3,
            "l": 1, "o": 1, "n": 1, "q": 10, "p": 3, "s": 1,
            "r": 1, "u": 1, "t": 1, "w": 4, "v": 4, "y": 4,
            "x": 8, "z": 10}

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
    '''Sort the rack alphebatically and return as a list. Make wildcards in the front.'''
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
        potential_rack = [[letter] + rack[1:] for letter in alphabet]
    elif rack[0] == '*' and rack[1] == '?':
        combs = list(combinations_with_replacement(alphabet, 2))
        potential_rack = [list(comb) + rack[2:] for comb in combs]
    else:
        potential_rack = rack
    return potential_rack

def match_words_w_wildcard(vocabs:list, rack_combs:list) -> list:
    '''
    Generate possible vocab list based on the rack. With wildcard
    '''
    matches = []
    used = [] # a list of used character that already filled with normal alphabet

    for rack_comb in rack_combs:
        for vocab in vocabs:
            valid = True
            char_freq = get_char_freq(vocab)
            for char in char_freq:
                if (char not in rack_comb) or (char_freq[char] > rack_comb.count(char)):
                    valid = False
            if valid:
                matches.append(vocab)
                used.append(rack_comb)
                continue
    matches_dict = {matches[i] : used[i] for i in range(len(matches))}
    return matches_dict

def score_words_w_wildcard(wildcard_dict, rack):
    '''
    This function calculate the scores of each word and return a 
    tuple of possible words with the score from high to low.
    '''
    scores = []

    char_freq_rack = get_char_freq(rack.lower())
    if ("*" in rack and not "?" in rack) or ("?" in rack and not "*" in rack):
        for word, letters in wildcard_dict.items():
            if letters[0] in word:
                wildcard = letters[0] # since the first letter is a substitute of the wildcard
                wildcard_val = score_dict.get(wildcard)
                score = sum([score_dict[character] for character in word]) - wildcard_val

                # Check if score can pump up by using original rack
                char_freq_word = get_char_freq(word.lower())
                use_original = True
                for char in char_freq_word:
                    if char_freq_word[char]>char_freq_rack.get(char,0):
                        use_original = False
                if use_original:
                    score = sum([score_dict[character] for character in word])
            else:
                score = sum([score_dict[character] for character in word])
            scores.append((score, word.upper()))
    elif ("*" in rack and "?" in rack) and (len(rack)==2):
        scores = [(0, vocab.upper()) for vocab in wildcard_dict if len(vocab) <= 2]
    elif ("*" in rack and "?" in rack) and (len(rack)>=2):
        for word, letters in wildcard_dict.items():
            if (letters[0] in word and not letters[1] in word) or (letters[1] in word and not letters[0] in word): # only one wildcard is used in a word
                wildcard = letters[0]
                wildcard_val = score_dict.get(wildcard)
                score = sum([score_dict[character] for character in word]) - wildcard_val
            elif (letters[0] in word) and (letters[1] in word): # when both wildcard substituted letter in a word from sowpods
                if letters[0] == letters[1]:
                    wildcard = letters[0]
                    wildcard_val = score_dict.get(wildcard)
                    score = sum([score_dict[character] for character in word]) - wildcard_val
                else:
                    wildcard1 = letters[0]
                    wildcard2 = letters[1]
                    wildcard_val_1 = score_dict.get(wildcard1)
                    wildcard_val_2 = score_dict.get(wildcard2)
                    score = sum([score_dict[character] for character in word]) - wildcard_val_1 - wildcard_val_2
            else:
                score = sum([score_dict[character] for character in word])
            scores.append((score, word.upper()))

    scores = sorted(scores, key = lambda x: (-x[0], x[1]))

    return (scores, len(scores))
