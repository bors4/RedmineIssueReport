from redminelib import Redmine
import yaml
from pathlib import Path
import os 

dir_path = os.path.dirname(os.path.realpath(__file__))

with open(dir_path + "\params.yaml", "r", encoding='utf8') as stream:
    try:
        params = yaml.safe_load(stream)
        REDMINE_URL = params["url"]
        API_KEY = params["key"]
        REDMINE_VERSION = params["version"]
        PROJECT_NAME = params["project"]
        SAVE_PATH = params["savepath"]
    except yaml.YAMLError as exc:
        print(exc)

redmine = Redmine(REDMINE_URL, key=API_KEY, raise_attr_exception=False, version=REDMINE_VERSION)

project = redmine.project.get(PROJECT_NAME)

issues = redmine.issue.filter(project_id=PROJECT_NAME, status_id='*')

issues.export('pdf', savepath=SAVE_PATH, columns='all')