

ship_distance = float(input("Enter the delivery distance in kilometers: "))
flat_rate_cost = 20.00
print("Flat rate shipping cost: $20.00")
dynamic_cost = 5 + 2 * ship_distance
print("Dynamic shipping cost:", dynamic_cost)
if flat_rate_cost < dynamic_cost:
    print("The flat rate shipping is cheaper.")
elif flat_rate_cost> dynamic_cost:
    print("The dynamic shipping is cheaper.")   
else:    print("Both shipping methods cost the same.")