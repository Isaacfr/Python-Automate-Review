import filenames, find_closest_project, compare_duplicates

find_closest_project.find_closest_project(filenames.source_file, filenames.source_sheet, filenames.lookup_file_path, filenames.lookup_table_file, filenames.lookup_full_path, filenames.lookup_table_sheet_name)

compare_duplicates.highlight_duplicates(filenames.source_file)

print("Found closest project and highlighted duplicates")