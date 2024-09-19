clear all; clc;
port_1_filename = '2024_02_21_remeasured_p1_shifted_sample_v2.xlsx';
port_2_filename = '2024_02_21_remeasured_p2_shifted_sample_v2.xlsx';

% Phi 0
% Call the function for Elevation e-theta (sheet index 2)
combine_and_save_data(port_1_filename, port_2_filename, 2);
% Call the function for Elevation e-phi (sheet index 3)
combine_and_save_data(port_1_filename, port_2_filename, 3);

% Phi 90
% Call the function for Azimuth e-theta (sheet index 4)
combine_and_save_data(port_1_filename, port_2_filename, 4);
% Call the function for Azimuth e-phi (sheet index 5)
combine_and_save_data(port_1_filename, port_2_filename, 5);

combine_csv_to_excel(port_1_filename);