"""
This program runs the scrabble cheater to generate a 
list of vocabularies and its score based on the rack
"""

from rack import check_rack, fill_wildcard
from wordscore import match_words_w_wildcard_2, score_word

def run_scrabble(rack: str):
    '''
    Input the rack as a string
    Check the rack input. Sort the rack. 
    Narrow Down possible vocabularies from sowpods. Generate strategies.
    '''
    # Check the input
    msg = check_rack(rack)
    if msg:
        return msg

    # Narrow down possible vocabs from sowpods
    with open("sowpods.txt","r", encoding="utf-8") as infile:
        raw_input = infile.readlines()
        data = [datum.strip('\n').lower() for datum in raw_input]
    trimmed_data = [item for item in data if len(item) <= len(rack)]

    # Fill the rack if there are wildcards
    possible_rack_combs = fill_wildcard(rack)

    # Generate strategies
    words = match_words_w_wildcard_2(trimmed_data, possible_rack_combs)
    res = score_word(words, rack)

    return res

if __name__ == '__main__':
    import pprint
    pprint.pprint(run_scrabble('PEN*?in'))
    
    
