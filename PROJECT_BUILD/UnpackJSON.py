import json
import logging.config

logging.config.fileConfig("../debug/logging.conf")
logger = logging.getLogger("UnpackJSON")


class UnpackJSON:
    def __init__(self, json_path):
        self.JSON_PATH = json_path

    def extract(self):
        logger.debug("begin JSON extraction")
        with open(self.JSON_PATH) as user_file:
            file_contents = user_file.read()
        logger.debug("loading JSON into dict object")
        parsed_json = json.loads(file_contents)  # loads JSON file
        logger.debug(parsed_json)
        logger.debug("extraction complete, returning extract to GENPPTX")
        return parsed_json
