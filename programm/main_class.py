from abc import ABC, abstractmethod
import pandas as pd

class BotSteps(ABC):

    def retrieve_steps_list(self):
        pass