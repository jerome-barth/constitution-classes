l = []  # liste vide
while True:
    try:
        val = float(input("Valeur à ajouter ou X pour terminer :"))
        l.append(val)
    except:
        break

# Traitement
if len(l) == 0:
    raise IndexError

l.sort()
if len(l) % 2:
    m = l[len(l) // 2]
else:
    m = (l[len(l) // 2 - 1] + l[len(l) // 2]) / 2

# Sortie
print(m)