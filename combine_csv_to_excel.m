function combine_csv_to_excel(port_filename)
    % Read data from CSV files
    elevation_theta_phase_data = csvread('Elevation_total_phase_e_theta.csv');
    elevation_theta_magnitude_data = csvread('Elevation_total_magnitude_e_theta.csv');
    elevation_phi_phase_data = csvread('Elevation_total_phase_e_phi.csv');
    elevation_phi_magnitude_data = csvread('Elevation_total_magnitude_e_phi.csv');

    azimuth_theta_phase_data = csvread('Azimuth_total_phase_e_theta.csv');
    azimuth_theta_magnitude_data = csvread('Azimuth_total_magnitude_e_theta.csv');
    azimuth_phi_phase_data = csvread('Azimuth_total_phase_e_phi.csv');
    azimuth_phi_magnitude_data = csvread('Azimuth_total_magnitude_e_phi.csv');

    % frequency points
    numbers_row = textread('numbers.txt', '%f');
    numbers_row = numbers_row';

    % Create the combined data array with the new row and empty rows
    empty_row_1 = NaN(1, size(elevation_theta_magnitude_data, 2));
    empty_row_2 = NaN(2, size(elevation_theta_magnitude_data, 2));
    elevation_theta_combined_data = [empty_row_1; numbers_row;  elevation_theta_magnitude_data; empty_row_2; numbers_row; elevation_theta_phase_data];
    elevation_phi_combined_data = [empty_row_1; numbers_row; elevation_phi_magnitude_data; empty_row_2; numbers_row; elevation_phi_phase_data];
    azimuth_theta_combined_data = [empty_row_1; numbers_row; azimuth_theta_magnitude_data; empty_row_2; numbers_row; azimuth_theta_phase_data];
    azimuth_phi_combined_data = [empty_row_1; numbers_row; azimuth_phi_magnitude_data; empty_row_2; numbers_row; azimuth_phi_phase_data];

    title_1 = {'T/x Magnitude (dB)'};
    title_2 = {'Theta Angle (Deg) \ Frequency (GHz)'};
    title_3 = {'T/x Phase (dB)'};
    numbers_1 = (-180:5:180)';
    numbers_2 = (0:5:360)';
    numbers_cell_1 = num2cell(numbers_1);
    numbers_cell_2 = num2cell(numbers_2);
    empty_row1 = {''};
    theta_cell_array_1 = [title_1; title_2; numbers_cell_1; empty_row1; title_3; title_2; numbers_cell_1];
    theta_cell_array_2 = [title_1; title_2; numbers_cell_2; empty_row1; title_3; title_2; numbers_cell_2];
    
    elevation_theta_combined_data = num2cell(elevation_theta_combined_data);
    elevation_phi_combined_data = num2cell(elevation_phi_combined_data);
    azimuth_theta_combined_data = num2cell(azimuth_theta_combined_data);
    azimuth_phi_combined_data = num2cell(azimuth_phi_combined_data);

    elevation_theta_combined_data_array = [theta_cell_array_1, elevation_theta_combined_data];
    elevation_phi_combined_data_array = [theta_cell_array_1, elevation_phi_combined_data];
    azimuth_theta_combined_data_array = [theta_cell_array_2, azimuth_theta_combined_data];
    azimuth_phi_combined_data_array = [theta_cell_array_2, azimuth_phi_combined_data];

    % Write data to an Excel file with multiple sheets
    copyfile('C:\Users\Jakub\OneDrive - Technological University Dublin\Jakub PhD\IEEE AWPL\MATLAB\Measurements\Combined_results_2024_02_21\2024_02_21_remeasured_p1_shifted_sample_v2.xlsx', 'C:\Users\Jakub\OneDrive - Technological University Dublin\Jakub PhD\IEEE AWPL\MATLAB\Measurements\Combined_results_2024_02_21\Combined_Data.xlsx');
    
    emptyData = {0};
    output_filename = 'Combined_Data.xlsx';
    elevation_theta_sheet_name = 'Elevation{Phi=0} E-Theta';
    elevation_phi_sheet_name = 'Elevation{Phi=0} E-Phi';
    azimuth_theta_sheet_name = 'Azimuth{Theta=90} E-Theta';
    azimuth_phi_sheet_name = 'Azimuth{Theta=90} E-Phi';
    empty_sheet = 'Return Loss'; 
    
    writecell(elevation_theta_combined_data_array, output_filename, 'Sheet', elevation_theta_sheet_name,'Range','A1'); % Starting from B1 to add one empty row
    writecell(elevation_phi_combined_data_array, output_filename, 'Sheet', elevation_phi_sheet_name, 'Range', 'A1'); % Starting from B1 to add one empty row
    writecell(azimuth_theta_combined_data_array, output_filename, 'Sheet', azimuth_theta_sheet_name,'Range','A1'); % Starting from B1 to add one empty row
    writecell(azimuth_phi_combined_data_array, output_filename, 'Sheet', azimuth_phi_sheet_name, 'Range', 'A1'); % Starting from B1 to add one empty row
    
end
