import os
import glob
import heapq
import codecs


def get_absentee_files(number):
    directory = 'C:' + os.sep + 'firefox_downloads'
    glob_match = directory + os.sep + 'Absentee*.xls'
    file_iterator = glob.iglob(glob_match)
    files = heapq.nlargest(number, file_iterator, key=os.path.getctime)
    return files


def get_email_file():
    directory = 'C:' + os.sep + 'firefox_downloads'
    glob_match = directory + os.sep + 'Jurisdiction*.xls'
    file_iterator = glob.iglob(glob_match)
    file = max(file_iterator, key=os.path.getctime)
    return file


def parse_file(filename):
    lines = []
    result = []
    with codecs.open(filename, 'r', 'utf-16') as file:
        # The first 125 lines are unnecessary header information
        lines = file.readlines()[125:]
    for line in lines:
        print(line)
        data = line.split('nobr>')[:-2]
        print(data)
        result.append(data)
    return result


def main():
    absentee_files = get_absentee_files(1)
    data = parse_file(absentee_files[0])
    print(absentee_files)
    email_file = get_email_file()
    print(email_file)


if __name__ == "__main__":
    main()

