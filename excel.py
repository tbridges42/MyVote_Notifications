import os
import glob
import heapq
import codecs
import csv


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


def parse_absentee_file(filename):
    result = []
    with codecs.open(filename, 'r', 'utf-16') as file:
        # The first 125 lines are unnecessary header information
        lines = file.readlines()[124:]
    for line in lines:
        data = [datum[:-2] for datum in line.replace("</td><td><nobr class='gridcellpadding'>", "").split('nobr>')
                if 'td' not in datum]
        if not data == []:
            result.append(data)
    return result[0:-2:2]


def parse_email_file(filename):
    result = dict()
    with codecs.open(filename, 'r', 'utf-16') as file:
        # The first 125 lines are unnecessary header information
        lines = file.readlines()[124:]
    for line in lines:
        data = [datum[:-2] for datum in line.split('nobr>')]
        if len(data) >= 6:
            result[data[1]] = data[4]
    return result


def add_emails(records, emails):
    for record in records:
        if record[10] in emails:
            record.append(emails[record[10]])
    return records


def print_records(records, filename):
    with open(filename, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(records)


def main(number):
    absentee_files = get_absentee_files(number)
    data = []
    for file in absentee_files:
        data += parse_absentee_file(file)
    email_file = get_email_file()
    emails = parse_email_file(email_file)
    data = add_emails(data, emails)
    print_records(data, 'C:\\firefox_downloads\\output.csv')


if __name__ == "__main__":
    main(1)

