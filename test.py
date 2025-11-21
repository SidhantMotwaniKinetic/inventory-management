def main():
    while True:
        scanned_value = input()
        if 'IN' in scanned_value and '-' in scanned_value:
            try:
                start_index = scanned_value.index('IN') + 2
                end_index = scanned_value.index('-', start_index)
                extracted_value = scanned_value[start_index:end_index]
                print(extracted_value)
            except ValueError:
                print("Error.") # Do nothing if pattern is not found
        else:
            print("Error.") # Do nothing if pattern not found

if __name__ == "__main__":
    main()