import time
import tensorflow as tf
import numpy as np

def compute_heavy_task_gpu():
    """A function that performs a more demanding numerical computation on the GPU."""
    matrix_size = 5000  # Larger matrix for a more intensive task
    A = tf.random.uniform((matrix_size, matrix_size), dtype=tf.float32)
    B = tf.random.uniform((matrix_size, matrix_size), dtype=tf.float32)
    tf.matmul(A, B)  # Matrix multiplication on the GPU

def main():
    num_tasks = 100  # Number of matrix multiplications

    print(f"Running test on the GPU...")
    start_time = time.time()

    for _ in range(num_tasks):
        compute_heavy_task_gpu()

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Time taken to complete the GPU test: {elapsed_time:.2f} seconds")
    
if __name__ == "__main__":
    main()  # Run the GPU test