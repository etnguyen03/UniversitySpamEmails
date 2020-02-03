import mailbox
import sys
import time

import tldextract
import datetime
import re
from dateutil import parser
import json
import csv
from email.header import decode_header


def mbox_reader(stream):
    """Read a non-ascii message from mailbox"""
    data = stream.read()
    text = data.decode(encoding="latin-1")
    return mailbox.mboxMessage(text)


def writeCollegeEmailListJSON(file, array):
    """
    Write JSON array out to file
    :param file: File path
    :param array: Array
    :return: None
    """
    with open(file, "w") as collegeListFile:
        collegeListFile.write(json.dumps(array, sort_keys=True, indent=4))


def main():
    timeBegin = time.perf_counter()
    messageParseList: [str, str, datetime.date] = []
    reEmail = re.compile("(?<=<).*(?=>)")
    for message in mailbox.mbox(sys.argv[1], factory=mbox_reader):
        # try:
        frm = message['from']

        if frm is None or frm == "error":
            continue

        subjectDecode = decode_header(message['subject'])
        subject = ""
        for string, charset in subjectDecode:
            if charset:
                subject += string.decode(charset)
            elif type(string) is str:
                subject += string
            else:
                subject += string.decode('utf-8')

        date = parser.parse(message['date'])

        if frm == "error" or frm is None:
            continue

        frmEmail = reEmail.findall(frm)

        if len(frmEmail) == 0:
            continue

        frmEmail = frmEmail[0]

        frmTLD = tldextract.extract(frmEmail)
        frmEmailTLD = frmTLD.domain + "." + frmTLD.suffix
        frmEmailTLD = frmEmailTLD.lower()

        messageParseList.append((frmEmailTLD, subject, date))

    # Read college list JSON file
    collegeListEmails = json.loads(open(sys.argv[2]).read())
    collegeListEmails = {k.lower(): v for k, v in collegeListEmails.items()}

    with open(sys.argv[3], "w") as csvfile:
        outCSV = csv.writer(csvfile, dialect='excel')
        outCSV.writerow(["College Name", "City, State", "Email Date", "Email Subject", "Subject Length"])
        for sender, subject, date in messageParseList:
            if sender not in collegeListEmails.keys():
                print("Sender", sender, "not found in dictionary")
                collegeName = input("Name of college, or -1 to skip: ")

                # Check for skip
                if collegeName == "-1":
                    print("SKIPPED")
                    collegeListEmails[sender] = "SKIP"
                    writeCollegeEmailListJSON(sys.argv[2], collegeListEmails)
                    continue

                collegeLocation = input("Location: ")

                collegeListEmails[sender] = [collegeName, collegeLocation]

                # Write JSON file
                writeCollegeEmailListJSON(sys.argv[2], collegeListEmails)

            if collegeListEmails[sender] == "SKIP":
                continue

            outCSV.writerow(
                [collegeListEmails[sender][0], collegeListEmails[sender][1], date.strftime("%Y-%m-%d"), subject,
                 len(subject)])

    # Output college list JSON file
    writeCollegeEmailListJSON(sys.argv[2], collegeListEmails)

    # Print stats
    print("Processed", len(messageParseList), "emails with", len(collegeListEmails), "colleges in", time.perf_counter()
          - timeBegin, "seconds")


if __name__ == '__main__':
    main()
