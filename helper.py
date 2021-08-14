#!/usr/bin/env python3
from difflib import SequenceMatcher
from typing import List

import pandas
from fuzzywuzzy import fuzz


def get_the_first_column_from_excel(filename: str) -> List[str]:
    """ reads an excel file, gets its first column and returns it as a list """
    excel_data_df = pandas.read_excel(filename, engine='openpyxl', header=None)

    # get the first column
    df = excel_data_df.iloc[:, 0]

    # drop the first and last rows
    df.drop(df.head(1).index, inplace=True)
    df.drop(df.tail(1).index, inplace=True)

    # convert the dataframe into a list and return it
    return df.to_list()


def get_name_from_email_address(email: str) -> str:
    # def get_name_from_email_address(email: str) -> List[str]:
    # strip the 4 symbol prefix and 12 symbol postfix
    # replace the '.' and '_' symbols between names with empty space
    # e.g 'vsp_surname.name@wh.domain.hu' -> 'surname name'
    return email[4:-12].replace('.', ' ').replace('_', ' ')


def flip_words(input_name):
    """ splits the string into a list and flips the list
        thus reversing the name. Returns a string.
    """
    name = input_name.split()
    name.reverse()
    return ' '.join(name)


def match(str1, str2):
    match = False
    for char1, char2 in zip(str1, str2):
        if char1 != char2:
            if match:
                return False
            else:
                match = True
    return match


def get_ratio(name_1, dict_name):
    """ gets 2 variants of a name by flipping the words.
        gets 2 ratios and returns the bigger ratio.
    """

    # get only the first half of the name
    common_word = False
    for idx, half in enumerate(name_1.split()):
        if half in dict_name and dict_name.find(half) != -1:
            common_word = half

    # if there is an IDENTICAL word - remove it
    # and compare the remaining words
    if common_word:
        name_1 = name_1.replace(common_word, '').strip()
        dict_name = dict_name.replace(common_word, '').strip()
        ratio = fuzz.partial_ratio(name_1, dict_name)
    else:
        name_2 = flip_words(name_1)
        ratio1 = fuzz.partial_ratio(name_1, dict_name)
        ratio2 = fuzz.partial_ratio(name_2, dict_name)
        ratio = max(ratio1, ratio2)

    return ratio


def compare_names(current_index, name, list_of_names, minimal_ratio):
    indexes_with_similarity = []
    for row, username in list_of_names.items():
        # do not compare with itself
        if row != current_index:
            ratio = get_ratio(name, username)
            if ratio >= minimal_ratio:
                indexes_with_similarity.append([row, username, ratio])
    # max_len = len(list_of_names)
    # if max_len > 0:
    #     for idx in range(2, max_len):
    #         ratio = get_ratio(name, list_of_names[idx])
    #         if ratio >= minimal_ratio:
    #             indexes_with_similarity.append([idx, list_of_names[idx], ratio])

    # returns all idx with similar names
    return indexes_with_similarity


def compare_name_with_timetable(row_index, name, timetable, minimal_ratio):
    indexes_with_similarity = []
    try:
        for timetable_idx, employee in enumerate(timetable, 2):
            if isinstance(employee[0], str) \
                    :
                username = employee[0]
                hours = employee[1]
                ratio = get_ratio(name, username.strip().lower())
                if ratio >= minimal_ratio:
                    indexes_with_similarity.append([row_index, name, timetable_idx, username, hours, ratio])
        return indexes_with_similarity
    except Exception as exception:
        raise Exception('compare_name_with_timetable', exception)


################
def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def ratios(name, employee_names):
    indexes_with_similarity = []
    max_len = len(employee_names) + 2
    for idx in range(2, max_len):
        ratio = similarity(name[0], employee_names[idx][0]) + \
                similarity(name[0], employee_names[idx][1]) + \
                similarity(name[1], employee_names[idx][0]) + \
                similarity(name[1], employee_names[idx][1])
        ratio /= 4
        # ratio = similarity(name, employee_names[idx])

        if ratio >= 0.52277:
            indexes_with_similarity.append(idx)

    return indexes_with_similarity


def yesno(question: str) -> bool:
    """Simple Yes/No Function."""
    prompt = f'{question} ? (y/n) default is "y": '
    ans = input(prompt).strip().lower() or 'y'
    if ans not in ['y', 'n']:
        print(f'{ans} is invalid, please try again...')
        return yesno(question)
    if ans == 'y':
        return True
    return False


def confirm_cell_write(ratio: int, index: int, text: str) -> bool:
    RESET_ALL = '\033[0m'
    LIGHT_BLUE = '\033[94m'
    LIGHT_YELLOW = '\033[93m'
    LIGHT_MAGENTA = '\033[95m'

    # some colored output added
    question = f'{LIGHT_YELLOW}ratio: {ratio} | {LIGHT_BLUE}index {index} |, {RESET_ALL}names: {LIGHT_MAGENTA}{text}.\n'
    question += f' {RESET_ALL}  -----------------------------'

    return yesno(question)
