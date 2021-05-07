import csv
import argparse
import time
import os
import zipfile
import json
import pathlib
import shutil
'''Create a CSV File with two columns:Name,MAC (if you use Excel, don't encode the CSV)'''

def main():
    parser = argparse.ArgumentParser(
        description='This scrip adds MAC Addresses to a pre-defined tag in a Ekahau project file.')
    parser.add_argument('file', metavar='esx_file', help='Ekahau project file')
    args = parser.parse_args()

    current_filename = pathlib.PurePath(args.file).stem
    working_directory = os.getcwd()
    print(args)

    with open('AP-to-MAC.csv') as file:
        csv_data = csv.DictReader(file, delimiter=';')

        with zipfile.ZipFile(args.file, 'r') as myzip:
            myzip.extractall(current_filename)

            # Load the accessPoints.json file into the accessPoints dictionary
            with myzip.open('accessPoints.json') as json_file:
                accessPoints = json.load(json_file)

            # Load the tagKeys.json file into the tagKeys dictionary
            with myzip.open('tagKeys.json') as json_file:
                tagKeys = json.load(json_file)

            for tag in tagKeys['tagKeys']:
                if tag['key'] == 'mac-address':
                    macaddress_id = tag['id']
            print(macaddress_id)
            for line in csv_data:
                for ap in accessPoints['accessPoints']:
                    if line['Name'] == ap['name']:
                        print('Match! ' + line['MAC'] + ' is assigned to ' + ap['name'])
                        ap['tags'].append(
                            {"tagKeyId": macaddress_id, "value": line['MAC']})

    with open(working_directory + '/' + current_filename + '/accessPoints.json', 'w') as file:
            json.dump(accessPoints, file, indent=4)

    new_filename = current_filename + '_modified'
    shutil.make_archive(new_filename, 'zip', current_filename)
    shutil.move(new_filename + '.zip', new_filename + '.esx')
    shutil.rmtree(current_filename)

if __name__ == "__main__":
    start_time = time.time()
    main()
    run_time = time.time() - start_time
    print("Time to run: %s sec" % round(run_time, 2))
