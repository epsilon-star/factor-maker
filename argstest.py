import sys

# sys.argv is a list of command-line arguments
print("All arguments:", sys.argv)

# First argument is always the script name
print("Script name:", sys.argv[0])

# Additional arguments
if len(sys.argv) > 1:
    print("First parameter:", sys.argv[1])
