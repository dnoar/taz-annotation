# -*- mode: Python; coding: utf-8 -*-

from classifier import Classifier
from scipy.special import logsumexp
from math import exp,log
from random import seed,shuffle

class MaxEnt(Classifier):

    
    def __init__(self, model={}):
        super(MaxEnt, self).__init__(model)

    def get_model(self): return self.meModel

    def set_model(self, model): self.meModel = model

    model = property(get_model, set_model)

    def train(self, instances, dev_instances=None, learning_rate=0.005, batch_size=30):
        """Construct a statistical model from labeled instances."""
        
        self.train_sgd(instances, dev_instances, learning_rate, batch_size)

    def train_sgd(self, instances, dev_instances, learning_rate, batch_size):
        """Train MaxEnt model with Mini-batch Stochastic Gradient 
        """
        
        #build feature vector and initialize parameters
        (featureVector,parameterDict,labels) = self.buildFeatureVector(instances)
        
        #initialize dev set instances as well
        for instance in dev_instances:
            self.buildDocFeatureVector(instance,featureVector)
        
        #save labels and feature vector; these won't change
        self.meModel['labels'] = labels
        self.meModel['featureVector'] = featureVector
        
        converged = 0
        bestResult = None
        
        #go until we haven't improved log likelihood
        while converged < 7:
            
            #chop up training set
            miniBatches = self.chopBatch(instances, batch_size, hash(bestResult))
            
            #compute the gradient for each minibatch, then update the parameters accordingly
            for miniBatch in miniBatches:
                gradient = self.computeGradient(miniBatch, featureVector, parameterDict,labels)
                for label in labels:
                    for i in gradient[label]:
                        parameterDict[label][i] += learning_rate * gradient[label][i]
            
            #find the log likelihood of the dev set given the new parameters
            result = self.logLikelihood(dev_instances, parameterDict)
            
            #save parameters if better than previous set
            if bestResult == None or result > bestResult:
                bestResult = result
                saveParameters = parameterDict.copy()
                converged = 0
            else:
                converged += 1
        
        #save final set of parameters
        self.meModel['parameterDict'] = saveParameters
        
    
    def buildFeatureVector(self, instances):
        '''Get all features from the training set and add indexes to each instance's feature vector'''
        featureList = []
        labels = []
        parameterDict = {}
        
        #Find all features for all instances, and add the index of the featureList to that instance
        for instance in instances:
            for feature in instance.features():
                
                try:
                    featureIndex = featureList.index(feature)
                except ValueError:
                    featureList.append(feature)
                    featureIndex = len(featureList)-1
                
                instance.addFeature(featureIndex)
                
            #also create list of labels
            if instance.label not in labels: labels.append(instance.label)
        
        #Create a parameter list for each feature for each class; initialize all to 0
        for label in labels:
            parameterDict[label] = [0] * len(featureList)
            
        return (featureList,parameterDict,labels)
            
    def computeGradient(self, instances, featureVector, parameterDict,labels):
        '''Compute the gradient for a set of instances given the current parameters'''
        
        #initialize gradient and its components
        gradientLeft = {}
        gradientRight = {}
        gradient = {}
        for label in labels:
            gradientLeft[label] = {}
            gradientRight[label] = {}
            gradient[label] = {}
        
        for instance in instances:
            instanceFeatures = instance.feature_vector
            
            #get the posterior for this instance for all labels
            instancePosterior = self.computeInstancePosteriors(instance,parameterDict,labels)
                
            for featureIndex in instanceFeatures:
                #at this point we have the index in the parameter vectors that this instance feature has
                
                #add one to the left side for the class that is this instance's label
                gradientLeft[instance.label][featureIndex] = gradientLeft[instance.label].setdefault(featureIndex,0) + 1
                
                #add the posterior to the right side for all labels
                for label in labels:
                    gradientRight[label][featureIndex] = gradientRight[label].setdefault(featureIndex,0) + instancePosterior[label]
        
        #compute the final gradient by subtracting right from left
        for label in labels:
            for i in gradientLeft[label]:
                gradient[label][i] = gradientLeft[label][i] - gradientRight[label][i]
             
        return gradient

    def computeInstancePosteriors(self, instance, parameterDict, labels):
        '''Find the posteriors for the instance for all possible labels given the current parameters'''
        instancePosteriors = {}
        allLabelVector = {}
        instanceFeatures = instance.feature_vector
        
        #add together all parameter weights for features in the instance
        for feature in instanceFeatures:
            for label in labels:
                allLabelVector[label] = allLabelVector.setdefault(label,0) + parameterDict[label][feature]
                
        #switch out the numerator to get posterior for all labels
        logSumExpValues = logsumexp(list(allLabelVector.values()))
        for label in labels:
            instancePosteriors[label] = exp(allLabelVector[label] - logSumExpValues)
            
        return instancePosteriors

    
    def computeLabelPosterior(self, instance, parameterDict, label):
        '''Find the posterior for the instance for one specific label'''
        labelSum = 0
        allLabelVector = {}
        instanceFeatures = instance.feature_vector
        
        for feature in instanceFeatures:
            for possLabel in parameterDict:
                allLabelVector[possLabel] = allLabelVector.setdefault(possLabel,0) + parameterDict[possLabel][feature]
            
        return exp(allLabelVector[label] - logsumexp(list(allLabelVector.values())))
     
    def chopBatch(self, instances, batch_size, seedVal):
        '''Chop a set of instances into batch_size pieces after randomizing based on the given seedVal'''
        
        miniBatches = []
        i = 0
        seed(seedVal)
        shuffle(instances)
        while i < len(instances):
            
            miniBatches.append(instances[i:(i+batch_size)])
            i += batch_size
        
        return miniBatches

    def logLikelihood(self, instances, parameterDict):
        '''Find the log likelihood of a batch of instances given the current parameters'''
        
        logLikelihood = 0
        for instance in instances:
        
            #find the posterior for the instance for its actual label
            instancePOS = self.computeLabelPosterior(instance, parameterDict, instance.label)
            
            #rare cases could create a math domain error
            if instancePOS > 0:
                logLikelihood += log(instancePOS)
        return logLikelihood
        
    def buildDocFeatureVector(self, instance, featureVector):
        '''Build an instance's feature_vector based on the featureVector of the training set'''
        
        #initialize it to nothing first
        instance.feature_vector = []
        
        #for each feature, see if it was in the features in the training set, and if so add its index to the instance's feature_vector
        for feature in instance.features():
            try:
                featureIndex = featureVector.index(feature)
            except ValueError:
                featureIndex = None
            if featureIndex != None:
                instance.addFeature(featureIndex)
    
    def classify(self, instance):
        '''Find the label with the highest posterior for the given instance'''
        goodLabel = ''
        goodLabelPos = 0
        labels = self.meModel['labels']
        
        #initialize instance to the model's feature vector
        self.buildDocFeatureVector(instance, self.meModel['featureVector'])
        
        #find the posterior for each label
        posteriors = self.computeInstancePosteriors(instance,self.meModel['parameterDict'],labels)
        
        #get the label with the highest posterior
        for label in labels:
            if posteriors[label] > goodLabelPos:
                goodLabelPos = posteriors[label]
                goodLabel = label
        
        return goodLabel
    
