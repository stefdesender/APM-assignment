import gurobipy as gp
from gurobipy import GRB

m = gp.Model("test")
x = m.addVar(vtype=GRB.BINARY, name="x")

m.setObjective(x, GRB.MAXIMIZE)
m.optimize()

print(f"Optimal value: {m.objVal}")
print(f"x: {x.x}")  

