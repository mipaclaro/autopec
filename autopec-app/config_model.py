# config_model.py
import configparser

class Config:
    def __init__(self, config_file='config.ini'):
        self.config = configparser.ConfigParser()
        self.config.read(config_file)

    @property
    def main_username(self):
        return self.config['credentials']['main_username']

    @property
    def main_password(self):
        return self.config['credentials']['main_password']

    @property
    def additional_username(self):
        return self.config['credentials']['additional_username']

    @property
    def additional_password(self):
        return self.config['credentials']['additional_password']

config = Config()