import FreeCAD
import Part
import random
import math
import numpy as np
import copy
from joblib import Parallel, delayed
from datetime import datetime

class GeneticAlgorithm:

    def __init__(self):
        super().__init__()
        self.vertices = []  
        self.population = []

    def load_vertices(self, file_path):
        shape = Part.Shape()
        shape.read(file_path)

        # Convert FreeCAD Vector objects to picklable tuples
        self.vertices = [(v.X, v.Y, v.Z) for v in shape.Vertexes]

    @staticmethod
    def fitness(chromosome, vertices):
        alpha, beta, gamma = [angle * (math.pi / 180) for angle in chromosome]

        # Setup rotation matrix
        R_x = np.array([[1, 0, 0], [0, math.cos(alpha), -math.sin(alpha)], [0, math.sin(alpha), math.cos(alpha)]])
        R_y = np.array([[math.cos(beta), 0, math.sin(beta)], [0, 1, 0], [-math.sin(beta), 0, math.cos(beta)]])
        R_z = np.array([[math.cos(gamma), -math.sin(gamma), 0], [math.sin(gamma), math.cos(gamma), 0], [0, 0, 1]])

        R = np.dot(np.dot(R_x, R_y), R_z)

        # Rotate vertices
        vertices = np.array(vertices)  # Convert list of tuples to NumPy array
        rotated_vertices = np.dot(vertices, R.T)

        # Calculate bounding box dimensions
        min_coords = np.min(rotated_vertices, axis=0)
        max_coords = np.max(rotated_vertices, axis=0)
        bbox = max_coords - min_coords
        volume = np.prod(bbox)

        return volume

    def parallel_fitness(self):
        fitness_scores = Parallel(n_jobs=-1)(
            delayed(self.fitness)(chromosome, self.vertices) for chromosome in self.population
        )
        return fitness_scores

    def selection(self):
        fitness_scores = self.parallel_fitness()
        fitness_volume = [[fitness_scores[i], i] for i in range(len(self.population))]
        fitness_volume.sort()

        if (len(self.population) // 2) % 2 == 0:
            half_size = (len(self.population)) // 2
        else:
            half_size = ((len(self.population)) // 2) + 1
            
        self.parents = [self.population[fitness_volume[i][1]] for i in range(half_size)]
       


    def crossover(self):
        self.children = []
        lst = list(range(0, len(self.parents)))

        while len(lst) >= 2:
            i = random.choice(lst)
            lst.remove(i)
            j = random.choice(lst)
            lst.remove(j)

            gene_idx = random.randint(0, 2)

            child1 = copy.deepcopy(self.parents[i])
            child2 = copy.deepcopy(self.parents[j])

            child1[gene_idx], child2[gene_idx] = child2[gene_idx], child1[gene_idx]

            self.children.append(child1)
            self.children.append(child2)

        # Select Elite
        if round((len(self.population) * 0.1)) % 2 == 0:
            elite = round((len(self.population) * 0.1))
        else:
            elite = round((len(self.population) * 0.1)) + 1

        for i in range(elite):
            self.children.append(self.parents[i])

    def mutation(self):
        for i in range(len(self.children)):
            n = 0.1
            if random.uniform(0, 1) <= n:
                gene_idx = random.randint(0, 2)
                if random.uniform(0, 1) < 0.5:
                    self.children[i][gene_idx] = min(self.children[i][gene_idx] + random.uniform(0,4), 90)
                else:
                    self.children[i][gene_idx] = max(self.children[i][gene_idx] - random.uniform(0,4), 0)

    def generate_population(self, pop_size):
        for _ in range(pop_size):
            gene = [round(random.uniform(0, 90), 2) for _ in range(3)]
            self.population.append(gene)

    def calc_bbox(self, rotation):
        alpha, beta, gamma = [angle * (math.pi / 180) for angle in rotation]

        R_x = np.array([[1, 0, 0], [0, math.cos(alpha), -math.sin(alpha)], [0, math.sin(alpha), math.cos(alpha)]])
        R_y = np.array([[math.cos(beta), 0, math.sin(beta)], [0, 1, 0], [-math.sin(beta), 0, math.cos(beta)]])
        R_z = np.array([[math.cos(gamma), -math.sin(gamma), 0], [math.sin(gamma), math.cos(gamma), 0], [0, 0, 1]])

        R = np.dot(np.dot(R_x, R_y), R_z)

        vertices = np.array(self.vertices)
        rotated_vertices = np.dot(vertices, R.T)

        min_coords = np.min(rotated_vertices, axis=0)
        max_coords = np.max(rotated_vertices, axis=0)
        bbox = max_coords - min_coords

        return bbox

    def opti_bbox(self, file_path, pop_size, itermax):
        self.load_vertices(file_path)
        self.generate_population(pop_size)

        func_eval = 0

        iteration = 0
        while iteration < itermax:
            iteration += 1
            func_eval += len(self.population)
            #print(f"Iteration {iteration}")

            self.selection()
            self.crossover()
            self.mutation()

            volume = []
            for i, child in enumerate(self.children):
                volume.append([self.fitness(child, self.vertices), i])
            volume.sort()
            #print(volume[0], len(self.children))
            self.population = self.children

        bbox = self.calc_bbox(self.children[volume[0][1]])
        #print(f"Optimum Bounding Box: {bbox}")
        print("Function Evaluations: ", func_eval)
        return bbox



    def bbox_plots(self, file_path, pop_size, itermax):
        self.load_vertices(file_path)
        self.generate_population(pop_size)

        func_eval = 0
        iter_volume = []

        iteration = 0
        while iteration < itermax:
            iteration += 1
            func_eval += len(self.population)
            #print(f"Iteration {iteration}")

            self.selection()
            self.crossover()
            self.mutation()

            volume = []
            for i, child in enumerate(self.children):
                volume.append([self.fitness(child, self.vertices), i])
            volume.sort()
            iter_volume.append(volume[0][0])
            
            print(volume[0], len(self.children))
            self.population = self.children
        #print("Best Volumes for each iteration: ", iter_volume)
        bbox = self.calc_bbox(self.children[volume[0][1]])
        print(f"Optimum Bounding Box: {bbox}")
        print(f"Optimum Volume: {volume[0][0]}")
        print("Function Evaluations: ", func_eval)
        return bbox, func_eval, iter_volume







if __name__ == '__main__':
    starttime = datetime.now()
    file_path = r'C:\Users\mail\OneDrive - Aalborg Universitet\9. Semester\P9 Code\StepFiles\BoundBoxTest.STEP'
    #file_path = r'C:\Cost_Calculation\StepFiles\BoundBoxTest.step'
    #file_path = 'C:\Cost_Calculation\StepFiles\Case1 - Part.step'
    ga = GeneticAlgorithm()
    bbox = ga.opti_bbox(file_path, pop_size=5000, itermax=15)
    print("Elapsed Time: ", datetime.now()-starttime)