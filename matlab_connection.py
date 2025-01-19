import matlab.engine

def run_macro_and_calculate_k_value(input_macro_path):
    # Start MATLAB engine
    eng = matlab.engine.start_matlab()

    try:
        isSuccess = eng.design_and_k_calculate(input_macro_path, nargout=1)  # Specify number of outputs
        k_value_file_path = "k_value.txt"
        eng.quit()

        return isSuccess, k_value_file_path
    except:
        print('Error in running the macro file for k calculation!')
        eng.quit()

        return None, None


def simulate_macro(input_macro_path):
    # Start MATLAB engine
    eng = matlab.engine.start_matlab()

    # Call a MATLAB script or function
    # Assuming `my_function.m` is in the MATLAB path and returns some values
    # Example: function [output1, output2] = my_function(input1, input2)

    try:
        isSuccess = eng.simulate(input_macro_path, nargout=1)  # Specify number of outputs
    
        eng.quit()
        return isSuccess
    except:
        print('Error in running the macro file for k calculation!')
        eng.quit()

        return None    
