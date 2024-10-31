n = input().split()
rows, columns = int(n[0]), int(n[1])
grid = []
for row in range(rows):
    column = input()
    grid.append(list(column))

rooms = []
for i in rows:
    for j in columns:
        if grid[i][j] == "." or grid[i][j] == "*":
            rooms.append()
