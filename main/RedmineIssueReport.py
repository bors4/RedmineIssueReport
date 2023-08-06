from redminelib import Redmine
import yaml
from pathlib import Path
import os
from datetime import datetime
import csv
import pandas as pd
dir_path = os.path.dirname(os.path.realpath(__file__))

with open(dir_path + "\params.yaml", "r", encoding="utf8") as stream:
    try:
        params = yaml.safe_load(stream)
        REDMINE_URL = params["url"]
        API_KEY = params["key"]
        REDMINE_VERSION = params["version"]
        PROJECT_NAME = params["project"]
        SAVE_PATH = params["savepath"]
        ENCODING_FN = params["encoding_fn"]
    except yaml.YAMLError as exc:
        print(exc)

redmine = Redmine(
    REDMINE_URL, key=API_KEY, raise_attr_exception=False, version=REDMINE_VERSION
)

print("Redmine server URL: " + redmine.url)

project = redmine.project.get(PROJECT_NAME)
print("Project name: " + project.name)

issues = redmine.issue.filter(project_id=PROJECT_NAME, status_id="*")

now = datetime.now()
dt_string = datetime.now().strftime("%d%m%Y%H%M%S")

print("Waiting for getting report")
fn = SAVE_PATH + "\\" + PROJECT_NAME + dt_string
issues.export("csv", savepath=SAVE_PATH, columns="all", filename=fn + ".csv")
print("File saved in:" + SAVE_PATH)
print("Convert from CSV format to *.xlsx format:\n")

# Open the input file with the original encoding
with open(fn + ".csv", "r", encoding="utf-8") as input_file:
    # Create a CSV reader object
    reader = csv.reader(input_file, dialect="excel", delimiter=";")
    ncol = len(next(reader))
    input_file.seek(0)
    included_cols = list(range(0,ncol))
    # Open the output file with the new encoding
    with open(
        fn + ENCODING_FN,
        "w",
        newline="",
        encoding="utf-16",
    ) as output_file:
        # Create a CSV writer object
        writer = csv.writer(output_file)
        # Write the rows to the output file
        print("Writing the rows to the output file\n")
        for row in reader:
            content = list(row[i] for i in included_cols)
            writer.writerow(content)

df = pd.read_csv(
    fn + ENCODING_FN,
    encoding="utf-16",
    sep=",",
    lineterminator="\n",
)

df.to_excel(fn + ".xlsx", index=None, header=True)

print("-----------------------------------------------------")
if os.path.isfile(fn + ENCODING_FN):
    os.remove(fn + ENCODING_FN)
    print("Temporary file " + fn + ENCODING_FN + " was deleted")
else:
    # If it fails, inform the user.
    print("Error:" + " file " + fn + ENCODING_FN + " not found")

if os.path.isfile(fn + ".csv"):
    os.remove(fn + ".csv")
    print("Temporary file " + fn + ".csv" + " was deleted")
else:
    # If it fails, inform the user.
    print("Error:" + " file " + fn + ".csv" + " not found")
