import random

# Given product data
'''products = [
    ("PI01", 55), ("PI02", 10), ("PI03", 10), ("PI04", 15), ("PI05", 5), ("PI06", 5), ("PI07", 10),
    ("PI08", 120), ("PI09", 20), ("PI10", 80), ("PI11", 40), ("PI12", 50), ("PI13", 60), ("PI14", 90),
    ("PI15", 30), ("PI16", 25), ("PI17", 70), ("PI18", 80), ("PI19", 150), ("PI20", 35), ("PI21", 100),
    ("PI22", 40), ("PI23", 50), ("PI24", 60), ("PI25", 300), ("PI26", 250), ("PI27", 75), ("PI28", 15),
    ("PI29", 80), ("PI30", 35), ("PI31", 350), ("PI32", 30), ("PI33", 45), ("PI34", 40), ("PI35", 25),
    ("PI36", 150), ("PI37", 70), ("PI38", 50), ("PI39", 120), ("PI40", 80), ("PI41", 60), ("PI42", 200),
    ("PI43", 500), ("PI44", 60), ("PI45", 300), ("PI46", 50), ("PI47", 150), ("PI48", 250), ("PI49", 25),
    ("PI50", 70), ("PI51", 150), ("PI52", 800), ("PI53", 120), ("PI54", 30), ("PI55", 100)
]

tp=[]

for i in range(26):
    # Randomly select a number from 1-10
    random_number = random.randint(1, 3)
    t=0
    # If the number is 3, select 3 random products and generate quantities
    selected_items = []
    selected_products = random.sample(products, random_number)
    for product_id, product_price in selected_products:
        quantity = random.randint(1, 4)
        flag=random.randint(0,2)
        if flag==1:
            quantity/=quantity
        total_amount = quantity * product_price
        t+=total_amount
        selected_items.append((product_id, quantity, total_amount))
    selected_items.append(t)
    tp.append(selected_items)
print(tp)'''
tp=[[('PI11', 4, 160), ('PI25', 3, 900), 1060], [('PI22', 2, 80), ('PI35', 2, 50), ('PI50', 1.0, 70.0), 200.0], [('PI43', 2, 1000), ('PI21', 3, 300), ('PI12', 3, 150), 1450], [('PI42', 1.0, 200.0), ('PI34', 1, 40), 240.0], [('PI04', 2, 30), ('PI52', 1, 800), 830], [('PI03', 4, 40), 40], [('PI29', 4, 320), 320], [('PI20', 3, 105), 105], [('PI04', 1.0, 15.0), 15.0], [('PI43', 4, 2000), ('PI41', 1.0, 60.0), ('PI55', 4, 400), 2460.0], [('PI33', 1.0, 45.0), ('PI17', 1, 70), ('PI15', 4, 120), 235.0], [('PI49', 1, 25), 25], [('PI38', 1, 50), ('PI37', 2, 140), 190], [('PI23', 3, 150), ('PI06', 1, 5), 155], [('PI06', 1, 5), ('PI15', 1, 30), 35], [('PI21', 1.0, 100.0), 100.0], [('PI25', 1.0, 300.0), ('PI54', 1, 30), ('PI28', 3, 45), 375.0], [('PI16', 1, 25), ('PI34', 4, 160), 185], [('PI54', 1, 30), 30], [('PI51', 4, 600), 600], [('PI44', 3, 180), ('PI37', 1.0, 70.0), 250.0], [('PI48', 4, 1000), 1000], [('PI14', 2, 180), ('PI43', 1, 500), 680], [('PI12', 3, 150), ('PI14', 2, 180), ('PI26', 1.0, 250.0), 580.0], [('PI03', 2, 20), ('PI39', 3, 360), ('PI24', 1.0, 60.0), 440.0], [('PI12', 1.0, 50.0), ('PI32', 3, 90), ('PI55', 4, 400), 540.0]]
for i in tp:
    print(i[-1])