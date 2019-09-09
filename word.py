# coding: utf-8
# input the number of word case. The default name of word case is word_9.*.xlsx. * is the input number
# The word case is a excel. Each row is a question entity.  The first column of each row is the question.
# And the second column should be the first solution. The third column is the second solution and so on.
# When the question has only one solution, the second column should be fill and the rest of the column should be left in
# blank.
# The system will ask question of each row with a random order, and it will ask every solution.
# When you are are going to input one of the solution, you can use help function.
# To input h or hh you will be return a help hint. At the first time you will get the first letter of the each word in
# the solution. For the second time you will get the second letter and so on.
# When you input something wrong accidentally and you want to remove that word from list, you can input r.
# It will remove the last word of the wrong list and you can answer the wrong word again.
# If you want to exit the system, press enter without any input five times continuously. The system will terminate.

import pyautogui
import xlrd
import numpy as np
from sys import exit


def generate_word_table():
    test_number = input("Word case：")
    print('Loading...')
    try:
        path_string = '/Users/suntuo/Desktop/未命名文件夹/写作/word_9.' + str(test_number) + '.xlsx'
        data = xlrd.open_workbook(path_string)
    except FileNotFoundError:
        print("\033[31;1mNo such file. Input again.\033[0m")
        return generate_word_table()
    else:
        table = data.sheets()[0]

        print('Succeeded...')
        print('Word list initiating...')

        table_to_list(table)

        return table


def table_to_list(table):
    word_list = []
    for i in np.arange(table.nrows):
        word_entity_row = table.row_values(i)
        # print(word_entity_row)
        word_entity_list = []
        for i in word_entity_row:
            if len(i) != 0:
                word_entity_list.append(i)
        word_list.append(word_entity_list)
        # print(len(word_entity_list))
    return word_list


def shuffled(word_list):
    shuffled_order = np.random.permutation(len(word_list))
    shuffled_list = []
    for i in shuffled_order:
        shuffled_list.append(word_list[i])
    start_str = 'There are ' + str(len(shuffled_list)) + ' word(s) in the case.'
    print(start_str)
    return shuffled_list


def test_print(word_for_recall, lenth, j):
    time = lenth[0] - 1 - j
    question_str = '[' + str(lenth[1])+ ']' + word_for_recall[0] + '(' + str(time) + ')' + ':'
    answer = input(question_str)
    if answer == '':
        return 'None'
    return answer


def answer_judge(answer, word):
    word_solution = word[1:]

    if answer == 'None':
        # print('blank')
        return 'blank', 0

    else:
        if answer == 'h' or answer == 'hh':
            return 'help', 0
        elif answer == 'r' or answer == 'rr':
            return 'recall', 0

        else:
            switch = 'wrong', 0
            for i in word_solution:
                if answer == i:
                    switch = 'right'
                    return switch, i
            return switch


def test_single(word, lenth, help_time, global_right, j):
    answer = test_print(word, lenth, j)
    # print(answer)
    switch_info = answer_judge(answer, word)
    # print(switch_info)
    switch = switch_info[0]
    # print(switch)
    if switch == 'help':
        return help_sys(word, lenth, help_time, global_right, j)

    elif switch == 'recall':
        return recall_current_sys(global_right), 0
    elif switch == 'blank':
        return 'blank', 0
    elif switch == 'wrong':
        return 'wrong', 0
    elif switch == 'right':
        return 'right', switch_info[1]


def recall_current_sys(global_right):
    if global_right == 0:
        return 'recall_current'
    else:
        return 'recall_last'


def recall_last_sys(word_list, wrong_list, i):
    try:
        wrong_list[-1][0] == word_list[i - 1][0]
    except IndexError:
        return False
    else:
        if len(wrong_list) != 0:
            if wrong_list[-1][0] == word_list[i - 1][0]:
                print("\033[34;1mRecall!\033[0m")
                return True
            else:
                return False
        else:
            return False


def help_sys(word, lenth, help_time, global_right, j):
    wordtohelp = word[1:]
    help_time += 1
    help_return = ''
    help_list = wordtohelp[0].split()
    for answer_word in help_list:
        help_word = ''
        # print(answer_word)
        if len(answer_word) - help_time >= 0:
            for i in np.arange(help_time):
                help_word += answer_word[i]
            for l in np.arange(len(answer_word) - help_time):
                help_word += '_'
        else:
            help_word = answer_word
        help_word += ' '
        help_return += help_word
    print(help_return)
    return test_single(word, lenth, help_time, global_right, j)


def test(word_list):
    i = 0
    blank_time = 0
    wrong_list = []
    # 单词词条
    while i <= len(word_list) - 1:
        # print(wrong_list)
        word = word_list[i]
        word_for_recall = word
        lenth = len(word_for_recall), len(word_list) - i
        # print(i)
        # 单个含义
        j = 0
        global_right = 1
        while j <= lenth[0] - 2:
            # print(j)
            # print(word_for_recall)
            if j == 0:
                word_pass = []

            help_time = 0
            solution_info = test_single(word, lenth, help_time, global_right, j)
            solution = solution_info[0]
            if solution == 'right':
                j += 1
                blank_time = 0
                word_pass.append(solution_info[1])
                word.remove(solution_info[1])
            elif solution == 'wrong':
                j += 1
                print("\033[31;1mWrong!\033[0m")
                blank_time = 0
                global_right = 0
                # print(word)
                # print(word)
            elif solution == 'recall_current':
                j = 0
                blank_time = 0
                global_right = 1
                if word_pass != 0:
                    for p in word_pass:
                        word.append(p)
            elif solution == 'recall_last':
                blank_time = 0

                if recall_last_sys(word_list, wrong_list, i):
                    global_right = -1
                    wrong_list.pop(-1)
                    break
                else:
                    print('Nothing to recall')

            elif solution == 'blank':
                blank_print = '?'
                for b in np.arange(blank_time):
                    blank_print += '?'
                blank_time = blank_exit(blank_time)
                print(blank_print)
        if global_right == 0:
            wrong_list.append(word_for_recall)
            i += 1
            Wrong_hint = ''
            for k in word_for_recall[1:]:
                Wrong_hint += k + '; '
            Wrong_hint = '\033[31;1m' + Wrong_hint + "\033[0m"
            print(Wrong_hint)
        elif global_right == 1:
            i += 1
        else:
            i -= 1

    return wrong_list


def blank_exit(blank_time):
    blank_time += 1
    if blank_time == 5:
        print("\033[31;1mExit detected!!!\033[0m")
        print("\033[32;1mExit detected!!!\033[0m")
        print("\033[34;1mExit detected!!!\033[0m")

        exit()
    return blank_time


def print_wrong(wrong_list):
    if len(wrong_list) == 0:
        print('Excellent!')
    else:
        wrong_question = []
        for i in np.arange(len(wrong_list)):
            wrong_question.append(wrong_list[i][0])
        max_len_str = max(wrong_question, key=len)
        for i in np.arange(len(wrong_list)):
            wrong_print = ''
            for wrong_word in wrong_list[i]:
                if len(wrong_word) != 0:
                    if len(wrong_print) == 0:
                        len_diff = len(max_len_str) - len(wrong_word)
                        for j in np.arange(2 * len_diff):
                            wrong_word += ' '
                        # print(len_diff)
                        wrong_word += ":"
                    wrong_print += wrong_word + ' '
            print(wrong_print)


def main():
    table = generate_word_table()
    word_list = shuffled(table_to_list(table))
    wrong_list = test(word_list)
    print_wrong(wrong_list)


if __name__ == "__main__":
    main()
