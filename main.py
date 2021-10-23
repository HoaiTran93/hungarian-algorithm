import sys
import matrixpriority as mp
import hungarian as ha

def main():
    print("Start!!!")
    dataPath = sys.argv[1]
    
    MatrixPriority = mp.MatrixPriority(dataPath)
    matrix_output = MatrixPriority.generate()

    hungarian_algorithm = ha.Hungarian(matrix_output)
    hungarian_algorithm.execute()

    print(hungarian_algorithm.toString())
    print("End!!!")

if __name__ == "__main__":
    main()