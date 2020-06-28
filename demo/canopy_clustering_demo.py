import math
import random
import numpy as np
import matplotlib.pyplot as plt


# 500 points 2 dimensions
dataset = np.random.rand(500, 2)


# main function
class Canopy:
    # init function for initializing the class variable
    def __init__(self, dataset):
        self.dataset = dataset
        self.t1 = 0
        self.t2 = 0

    # evaluate the distance between two vectors
    @staticmethod
    def euclideanDistance(vec1, vec2):
        return math.sqrt(((vec1 - vec2) ** 2).sum())

    # clustering method
    def clustering(self):

        if self.t1 == 0:
            print('Please set the threshold.')
        else:
            # this dictionary container will persist the result
            canopies = []
            while len(self.dataset) != 0:
                rand_index = self.getRandIndex()
                # get a point randomly as p
                current_center = self.dataset[rand_index]
                # initialize a container current_center_list for class canopy
                current_center_list = []
                # initialize a container delete_list for deleting points
                delete_list = []
                self.dataset = np.delete(
                    self.dataset, rand_index, 0)
                for datum_j in range(len(self.dataset)):
                    datum = self.dataset[datum_j]
                    distance = self.euclideanDistance(
                        current_center, datum)
                    if distance < self.t1:
                        # if distance < t1,add the point to current_center_list
                        current_center_list.append(datum)
                    if distance < self.t2:
                        # if distance < t2,add the point to delete_list
                        delete_list.append(datum_j)
                # according to the condition of the delete_list,delete the elements from the dataset
                self.dataset = np.delete(self.dataset, delete_list, 0)
                canopies.append((current_center, current_center_list))
            return canopies

    def setThreshold(self, t1, t2):
        if t1 > t2:
            self.t1 = t1
            self.t2 = t2
        else:
            print('t1 needs to be larger than t2!')

    def getRandIndex(self):
        return random.randint(0, len(self.dataset) - 1)


# Exhibition
def showCanopy(canopies, dataset, t1, t2):

    fig = plt.figure()
    sc = fig.add_subplot(111)
    colors = ['brown', 'green', 'blue', 'y', 'r', 'tan', 'dodgerblue', 'deeppink', 'orangered', 'peru', 'blue',
              'y', 'r', 'gold', 'dimgray', 'darkorange', 'peru', 'blue', 'y', 'r', 'cyan', 'tan', 'orchid',
              'peru', 'blue', 'y', 'r', 'sienna']
    markers = ['*', 'h', 'H', '+', 'o', '1', '2', '3', ',', 'v', 'H', '+', '1', '2', '^',  '<', '>', '.', '4', 'H', '+', '1', '2', 's', 'p', 'x', 'D', 'd', '|', '_']
    for i in range(len(canopies)):
        canopy = canopies[i]
        center = canopy[0]
        components = canopy[1]
        sc.plot(center[0], center[1], marker=markers[i],
                color=colors[i], markersize=10)
        t1_circle = plt.Circle(
            xy=(center[0], center[1]), radius=t1, color='dodgerblue', fill=False)
        t2_circle = plt.Circle(
            xy=(center[0], center[1]), radius=t2, color='skyblue', alpha=0.2)
        sc.add_artist(t1_circle)
        sc.add_artist(t2_circle)
        for component in components:
            sc.plot(component[0], component[1],
                    marker=markers[i], color=colors[i], markersize=1.5)

    maxvalue = np.amax(dataset)
    minvalue = np.amin(dataset)
    plt.xlim(minvalue - t1, maxvalue + t1)
    plt.ylim(minvalue - t1, maxvalue + t1)
    plt.show()


def main():

    t1 = 0.6
    t2 = 0.4
    gc = Canopy(dataset)
    gc.setThreshold(t1, t2)
    canopies = gc.clustering()
    print('Get %s initial centers.' % len(canopies))
    showCanopy(canopies, dataset, t1, t2)


if __name__ == '__main__':
    main()
