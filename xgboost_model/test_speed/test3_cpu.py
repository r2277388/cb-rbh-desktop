import time
from multiprocessing import Pool, cpu_count
import numpy as np

def compute_heavy_task(_):
    """A function that performs a more demanding numerical computation."""
    matrix_size = 5000  # Larger matrix for a more intensive task
    A = np.random.rand(matrix_size, matrix_size)
    B = np.random.rand(matrix_size, matrix_size)
    np.dot(A, B)  # Matrix multiplication

def main():
    num_processes = cpu_count()  # Use all available cores
    num_tasks = 100  # Increase the number of tasks for more computation

    print(f"Running test with {num_processes} processes...")
    start_time = time.time()

    with Pool(num_processes) as pool:
        pool.map(compute_heavy_task, range(num_tasks))

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Time taken to complete the test: {elapsed_time:.2f} seconds")
    
if __name__ == "__main__":
    main() 