import time
from multiprocessing import Pool, cpu_count
import numpy as np

def compute_heavy_task(_):
    """A function that performs a heavy numerical computation to simulate CPU load."""
    matrix_size = 1000
    A = np.random.rand(matrix_size, matrix_size)
    B = np.random.rand(matrix_size, matrix_size)
    np.dot(A, B)  # Matrix multiplication

def main():
    num_processes = cpu_count()  # Automatically use all available cores
    num_tasks = 50  # Number of tasks to process (adjust if necessary)

    print(f"Running test with {num_processes} processes...")
    start_time = time.time()

    with Pool(num_processes) as pool:
        pool.map(compute_heavy_task, range(num_tasks))

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Time taken to complete the test: {elapsed_time:.2f} seconds")
    
if __name__ == "__main__":
    main()  