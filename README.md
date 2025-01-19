# antenna-design-FORA
This is a repository documenting the application for dipole FORA antenna design and deployement to CST environment.



- **Modes of Operation:**
  - The developed code operates in two modes controlled by the global variable `GUI_MODE` in `config.py`:
    1. **Exploratory Mode:**
       - Enables users to define parameter ranges.
       - Iterates through the parameter space in predefined steps to evaluate performance criteria.
       - Saves parameter configurations meeting the criteria and generates performance assessment plots saved locally for analysis.
    2. **Direct Input Mode:**
       - Provides a graphical user interface (GUI) for direct parameter input.
       - Automates the simulation based on the submitted inputs.

- **Behavior Based on `GUI_MODE`:**
  - If `GUI_MODE` is `True`, the program activates Direct Input Mode with a GUI.
  - If `GUI_MODE` is `False`, the program performs iterative simulations over multiple parameter settings.

- **Parameter Configuration:**
  - All configurable parameters are listed in the article as a table.
  - Currently, there is no GUI available for the iterative loop scenario; development is in progress and will be shared very soon!

- **Iterative Parameter Search:**
  - Parameter ranges for the iterative simulation can be defined in the bottommost code snippet in `app.py`.
  - Users can add loops for the corresponding parameters to iteratively run simulations and retrieve results for each configuration set.

- **Instructions for `Calculate-port-extension-coefficient-MWS.mcr`:**
  - Replace the existing `Calculate-port-extension-coefficient-MWS.mcr` file in `Solver\Ports\` within the CST Studio program files with the updated file from the repository.
  - The updated file writes the calculated extension coefficient to `k_value.txt` in the current working directory of `app.py`.
  - Update the file path for saving `k_value.txt` in the `.mcr` file to match the `app.py` folder.

- **About `k_value.txt`:**
  - This file contains the calculated port extension coefficient value.
  - The example file in the repository serves as a reference.

- **Usage of MATLAB Files:**
  - **`design_and_k_calculate.m`:**  
    - Runs simulations with the intended geometry.  
    - Calculates and writes the port extension coefficient.  
  - **`simulate.m`:**  
    - Executes simulation commands.
  - Ensure the mathematical definitions in both `.m` files are reviewed for accuracy.
 
Output:
![dipole_antenna](https://github.com/user-attachments/assets/cd267d7e-32a2-4da1-9fcb-15cc085679e5)

