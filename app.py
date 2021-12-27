import pandas as pd
import math

def main():
    df = pd.read_excel('grades.xlsx', usecols=['Grade', '# of Units'])

    point12_conversion = {
        'A+': 12,
        'A': 11,
        'A-': 10,
        'B+': 9,
        'B': 8,
        'B-': 7,
        'C+': 6,
        'C': 5,
        'C-': 4,
        'D+': 3,
        'D': 2,
        'D-': 1,
        'F': 0
    }

    point4_conversion = {
        12: 4,
        11: 3.9,
        10: 3.7,
        9: 3.3,
        8: 3,
        7: 2.7,
        6: 2.3,
        5: 2,
        4: 1.7,
        3: 1.3,
        2: 1,
        1: 0.7,
        0: 0
    }

    total_units = 0
    total_gradepoints = 0
    for index, row in df.iterrows():
        units = row[1]
        grade = row[0]

        total_units += units
        try:
            total_gradepoints += point12_conversion[grade] * units
        except:
            print('Invalid grade letter')
            return

    cumulative12 = round(total_gradepoints / total_units, 2)
    
    if isinstance(cumulative12, float):
        lower = math.floor(cumulative12)
        upper = lower + 1

        print(upper, lower)
        cumulative4 = (((point4_conversion[upper] - point4_conversion[lower]) / (upper - lower)) * (cumulative12 - lower)) + point4_conversion[lower]
    else:
        cumulative4 = point4_conversion[cumulative12]

    cumulative4 = round(cumulative4, 4)

    print("12 point cumulative GPA: " + str(cumulative12))
    print("4 point cumulative GPA: " + str(cumulative4))
    return
    
    

if __name__ == '__main__':
    main()
