import time
import numpy as np

def single_core_test():
    matrix_size = 1000
    A = np.random.rand(matrix_size, matrix_size)
    B = np.random.rand(matrix_size, matrix_size)
    start_time = time.time()
    np.dot(A, B)
    end_time = time.time()
    return end_time - start_time

def main():
    elapsed_time = single_core_test()
    print(f"Single-core time taken: {elapsed_time:.2f} seconds")
    
if __name__ == "__main__":
    main()

