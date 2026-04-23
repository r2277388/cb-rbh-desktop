import test1_cpu
import test2_cpu
import test3_cpu
import test4_gpu

print("Running test1_cpu...")
test1_cpu.main()  # Call the main function in each script

print("Running test2_cpu...")
test2_cpu.main()

print("Running test3_cpu...")
test3_cpu.main()

print("Running test4_gpu...")
test4_gpu.main()

print("All scripts completed.")