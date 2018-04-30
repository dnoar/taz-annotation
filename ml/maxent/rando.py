# -*- mode: Python; coding: utf-8 -*-

from classifier import Classifier
from scipy.special import logsumexp
from math import exp,log
import random

class Rando(Classifier):

    
    def __init__(self, label_list, model={}):
        self.label_list = label_list
        super(Rando, self).__init__(model)

    def get_model(self): return self.meModel

    def set_model(self, model): self.meModel = model

    model = property(get_model, set_model)

    def train(self, instances, dev_instances=None):
        """Construct a statistical model from labeled instances."""
        pass
    
    def classify(self, instance):
        '''Find the label with the highest posterior for the given instance'''
        return self.label_list[random.randint(0,len(self.label_list)-1)]
    
