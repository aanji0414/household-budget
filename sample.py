def calculate_average(numbers):
    total = 0
    for n in numbers:
        total += n
    return total / len(numbers)

def find_max(numbers):
    max_val = numbers[0]
    for n in numbers:
        if n > max_val:
            max_val = n
    return max_val

scores = [85, 92, 78, 95, 88]
print("평균:", calculate_average(scores))
print("최고점:", find_max(scores))
