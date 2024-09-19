function combine_and_save_data(port_1_filename, port_2_filename, sheet_index)
    % Elevation e-theta and e-phi
    %clc;clearvars -except port_1_filename port_2_filename sheet_index;

    % Read data for port 1 and port 2
    % When doing scanning from -10 to 10 for elevation and 0 to 10 for
    % azimuth, we need to use this script 

    if sheet_index > 3 % -------> Azimuth
        P1_data_mag = readmatrix(port_1_filename, 'Sheet', sheet_index, 'Range', 'B3:SH75');
        P1_data_phase = readmatrix(port_1_filename, 'Sheet', sheet_index, 'Range', 'B79:SH151');
        P2_data_mag = readmatrix(port_2_filename, 'Sheet', sheet_index, 'Range', 'B3:SH75');
        P2_data_phase = readmatrix(port_2_filename, 'Sheet', sheet_index, 'Range', 'B79:SH151');
    else % -------> Elevation
        P1_data_mag = readmatrix(port_1_filename, 'Sheet', sheet_index, 'Range', 'B3:SH75');
        P1_data_phase = readmatrix(port_1_filename, 'Sheet', sheet_index, 'Range', 'B79:SH151');
        P2_data_mag = readmatrix(port_2_filename, 'Sheet', sheet_index, 'Range', 'B3:SH75');
        P2_data_phase = readmatrix(port_2_filename, 'Sheet', sheet_index, 'Range', 'B79:SH151');
    end

    % Constants
    rows = size(P1_data_mag, 1);
    columns = size(P1_data_mag, 2);

    % Preallocate memory for result arrays
    total_phase = zeros(rows, columns);
    total_magnitude = zeros(rows, columns);
    
    % Loop through the data
    for column = 1:columns
        for row = 1:rows
            % Port 1
            magnitude_p1 = db2mag(P1_data_mag(row, column));
            magnitude_p1 = magnitude_p1 / sqrt(2);
            phase_p1 = deg2rad(P1_data_phase(row, column));
            [real_p1, imaginary_p1] = pol2cart(phase_p1, magnitude_p1);

            % Port 2
            magnitude_p2 = db2mag(P2_data_mag(row, column));
            magnitude_p2 = magnitude_p2 / sqrt(2);
            phase_p2 = P2_data_phase(row,column) + 90;
            if phase_p2 > 180
                phase_p2 = -180 + (phase_p2-180);
            elseif phase_p2 < -180
                phase_p2 = 180 + (phase_p2-180);
            end
            phase_p2 = deg2rad(phase_p2);
            [real_p2, imaginary_p2] = pol2cart(phase_p2, magnitude_p2);

            % Addition
            total_real = real_p1 + real_p2;
            total_imaginary = imaginary_p1 + imaginary_p2;
            [total_phase(row, column), total_magnitude(row, column)] = cart2pol(total_real, total_imaginary);
            total_phase(row, column) = rad2deg(total_phase(row, column)); % convert rad to degrees
            total_magnitude(row, column) = mag2db(total_magnitude(row, column));
        end
    end

    % Export to CSV files
    if sheet_index == 2
        csvwrite('Elevation_total_phase_e_theta.csv', total_phase);
        csvwrite('Elevation_total_magnitude_e_theta.csv', total_magnitude);
    elseif sheet_index == 3
        csvwrite('Elevation_total_phase_e_phi.csv', total_phase);
        csvwrite('Elevation_total_magnitude_e_phi.csv', total_magnitude);
    elseif sheet_index == 4
        csvwrite('Azimuth_total_phase_e_theta.csv', total_phase);
        csvwrite('Azimuth_total_magnitude_e_theta.csv', total_magnitude);
    elseif sheet_index == 5
        csvwrite('Azimuth_total_phase_e_phi.csv', total_phase);
        csvwrite('Azimuth_total_magnitude_e_phi.csv', total_magnitude);
    end  
end
