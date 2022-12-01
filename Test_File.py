import random

# List of most common Powerball numbers based on the last 7 years
lucky_numbers = (4,6,10,18,21,23,24,32,61,63,69)

# Generate 5 set of 6 non-repeated random numbers from the list with the last number smaller than 26 being the Powerball number
for i in range(5):
    n = 1
    selected_numbers = []
    while n < 7:
        lucky_number = random.randrange(0,len(lucky_numbers))
        # If the random number already selected, keep rolling, we don't want any repeated numbers
        while lucky_numbers[lucky_number] in selected_numbers:
            # Keep rolling if the selected number for the Powerball number is bigger than 26
            while n == 6 and lucky_numbers[lucky_number] > 26:
                lucky_number = random.randrange(0,len(lucky_numbers))
            else:
                lucky_number = random.randrange(0,len(lucky_numbers))
        else:
            # Keep rolling if the selected number for the Powerball number is bigger than 26
            while n == 6 and lucky_numbers[lucky_number] > 26:
                lucky_number = random.randrange(0,len(lucky_numbers))
            else:
                selected_numbers.append(lucky_numbers[lucky_number])
        n += 1
    print(selected_numbers)