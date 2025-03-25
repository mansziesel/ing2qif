#!/usr/bin/python3
# (C) 2014, Marijn Vriens <marijn@metronomo.cl>
# GNU General Public License, version 3 or any later version
# Small additions by Steven Jelsma and Mans Ziesel

# Documents
# https://github.com/Gnucash/gnucash/blob/master/src/import-export/qif-imp/file-format.txt
# https://en.wikipedia.org/wiki/Quicken_Interchange_Format

import csv
import argparse
import sys


class Entry:
    """
    I represent one entry.
    """

    def __init__(self, data):
        self._data = data
        self._clean_up()

    def _clean_up(self):
        self._data['amount'] = self._data['Bedrag (EUR)'].replace(',', '.')

    def keys(self):
        return self._data.keys()

    def __getattr__(self, item):
        try:
            return self._data[item]
        except KeyError:
            raise AttributeError(f"'Entry' object has no attribute '{item}'")

    def __getitem__(self, item):
        if item not in self._data:
            raise KeyError(f"Key '{item}' not found in entry data. Available keys: {self._data.keys()}")
        return self._data[item]


class CsvEntries:
    def __init__(self, filedescriptor):
        self._entries = csv.DictReader(filedescriptor)

    def __iter__(self):
        return map(Entry, self._entries)


class QifEntries:
    def __init__(self):
        self._entries = []

    def add_entry(self, entry):
        """
        Add an entry to the list of entries in the statement.
        :param entry: A dictionary where each key is one of the keys of the statement.
        :return: Nothing.
        """
        self._entries.append(QifEntry(entry))

    def serialize(self):
        """
        Turn all the entries into a string
        :return: a string with all the entries.
        """
        data = ["!Type:Bank"]
        for e in self._entries:
            data.append(e.serialize())
        return "\n".join(data)


class QifEntry:
    def __init__(self, entry):
        self._entry = entry
        self._data = []
        self.processing(self._data)

    def processing(self, data):
        # Add date
        data.append(f"D{self._date_format(self._entry['Datum'])}")  # Use the formatted date

        # Add amount (formatted)
        data.append(f"T{self._amount_format()}")

        # Add amount (formatted)
        data.append(f"U{self._amount_format()}")

        # Add payee (P line with the contents of 'Naam / Omschrijving')
        data.append(f"P{self._entry['Naam / Omschrijving']}")

        # Add memo
        data.append(f"M{self._memo()}")

        # Add Cc line
        data.append("Cc")  # Always add "Cc" line

        # Add transaction type if available
        # if self._entry_type():
        #    data.append(f'N{self._entry_type()}')

        # Add N line
        data.append("N")  # Always add "Cc" line

        # End of entry
        data.append("^")

    def serialize(self):
        """
        Turn the QifEntry into a String.
        :return: a string
        """
        return "\n".join(self._data)

    def _memo_geldautomaat(self, mededelingen, omschrijving):
        if omschrijving.startswith('ING>') or \
                omschrijving.startswith('ING BANK>') or \
                omschrijving.startswith('OPL. CHIPKNIP'):
            return omschrijving
        else:
            return mededelingen[:32]

    def _memo_incasso(self, mededelingen, omschrijving):
        if omschrijving.startswith('SEPA Incasso') or mededelingen.startswith('SEPA Incasso'):
            try:
                s = mededelingen.index('Naam: ') + 6
                e = mededelingen.index('Kenmerk: ')
                return mededelingen[s:e]
            except ValueError:
                raise Exception(mededelingen, omschrijving)

    def _memo_internetbankieren(self, mededelingen, omschrijving):
        try:
            s = mededelingen.index('Naam: ') + 6
            if "Omschrijving:" in mededelingen:
                e = mededelingen.index('Omschrijving: ')
            else:
                e = mededelingen.index('IBAN: ')
            return mededelingen[s:e]
        except ValueError:
            return None

    def _memo_diversen(self, mededelingen, omschrijving):
        return mededelingen[:64]

    def _memo_verzamelbetaling(self, mededelingen, omschrijving):
        if 'Naam: ' in mededelingen:
            s = mededelingen.index('Naam: ') + 6
            e = mededelingen.index('Kenmerk: ')
            return mededelingen[s:e]

    def _memo(self):
        """
        Decide what the memo field should be. Try to keep it as sane as possible. If unknown type, include all data.
        :return: the memo field.
        """
        mutatie_soort = self._entry['Mutatiesoort']
        mededelingen = self._entry['Mededelingen']
        omschrijving = self._entry['Naam / Omschrijving']

        memo = None
        try:
            memo_method = {  # Depending on the mutatie_soort, switch memo generation method.
                'Diversen': self._memo_diversen,
                'Betaalautomaat': self._memo_geldautomaat,
                'Geldautomaat': self._memo_geldautomaat,
                'Incasso': self._memo_incasso,
                'Internetbankieren': self._memo_internetbankieren,
                'Overschrijving': self._memo_internetbankieren,
                'Verzamelbetaling': self._memo_verzamelbetaling,
            }[mutatie_soort]
            memo = memo_method(mededelingen, omschrijving)
        except KeyError:
            pass
        finally:
            if memo is None:
                # The default memo value. All the text.
                memo = f"{self._entry['Mededelingen']} {self._entry['Naam / Omschrijving']}"
        if self._entry_type():
            return f"{self._entry_type()} {memo}"
        return memo.strip()

    def _amount_format(self):
        if self._entry['Af Bij'] == 'Bij':
            return "+" + self._entry['amount']
        else:
            return "-" + self._entry['amount']
        
    def _date_format(self, date_str):
        """Convert date from YYYYMMDD to DD/MM/YYYY format."""
        return f"{date_str[6:8]}/{date_str[4:6]}/{date_str[0:4]}"

    def _entry_type(self):
        """
        Detect the type of entry.
        :return:
        """
        try:
            return {
                'Geldautomaat': "ATM",
                'Internetbankieren': "Transfer",
                'Incasso': 'Transfer',
                'Verzamelbetaling': 'Transfer',
                'Betaalautomaat': "ATM",
                'Storting': 'Deposit',
            }[self._entry['Mutatiesoort']]
        except KeyError:
            return None


def main(filedescriptor, start, number, input_filename):
    qif = QifEntries()
    c = 0
    for entry in CsvEntries(filedescriptor):
        c += 1
        if c >= start:
            qif.add_entry(entry)
            if number and c > start + number - 2:
                break

    # Generate output filename by replacing .csv with .qif
    if input_filename.endswith(".csv"):
        output_file = input_filename[:-4] + ".qif"
    else:
        output_file = input_filename + ".qif"

    # Write the serialized QIF data to the output file with UTF-8 encoding
    with open(output_file, 'w', encoding='utf-8') as out_fd:
        out_fd.write(qif.serialize())

    print(f"QIF data written to {output_file}")


def parse_cmdline():
    # If a file is dragged and dropped, sys.argv will contain the path as an argument
    if len(sys.argv) > 1:
        # sys.argv[0] is the script itself, sys.argv[1] will be the dragged CSV file
        csvfile = sys.argv[1]
        start = 0
        number = None
    else:
        # Fall back to argparse if no file is dragged
        parser = argparse.ArgumentParser(description="Convert ING banking statements in CSV format to QIF file for GnuCash.")
        parser.add_argument("csvfile", metavar="CSV_FILE", help="The CSV file with banking statements.")
        parser.add_argument("--start", type=int, metavar="NUMBER", default=0, help="The statement you want to start conversion at.")
        parser.add_argument("--number", type=int, metavar="NUMBER", help="The number of statements to convert")
        args = parser.parse_args()
        csvfile = args.csvfile
        start = args.start
        number = args.number

    return csvfile, start, number


if __name__ == '__main__':
    # Get the file path and other parameters
    csvfile, start, number = parse_cmdline()
    # Open the file in UTF-8 mode and run the conversion
    with open(csvfile, 'r', encoding='utf-8') as fd:
        main(fd, start, number, csvfile)
