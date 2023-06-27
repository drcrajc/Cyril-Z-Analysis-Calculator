import pandas as pd
import math
import matplotlib.pyplot as plt

# Read data from Excel file
df = pd.read_excel('data.xlsx', engine='openpyxl')
frequency = df['f'].tolist()
Zm = df['Zm'].tolist()
Zp = df['Zp'].tolist()
Cap = 10  # Example capacitance value

# Calculate the data columns
Z_prime = [Z * math.cos(math.radians(p)) for Z, p in zip(Zm, Zp)]
Z_double_prime = [-Z * math.sin(math.radians(p)) for Z, p in zip(Zm, Zp)]
Z_double_prime_max = max(Z_double_prime)
Z_double_prime_normalized = [Z / Z_double_prime_max for Z in Z_double_prime]
e_prime = [Z / ((2 * math.pi * f) * Cap * ((d ** 2) + (e ** 2))) for f, d, e in zip(frequency, Z_prime, Z_double_prime)]
e_double_prime = [Z / ((2 * math.pi * f) * Cap * ((d ** 2) + (e ** 2))) for f, d, e in zip(frequency, Z_double_prime, Z_prime)]
e_star = [math.sqrt((ep ** 2) + (epp ** 2)) for ep, epp in zip(e_prime, e_double_prime)]
tan_delta = [hp / ep for hp, ep in zip(e_double_prime, e_prime)]
M_star = [1 / ep for ep in e_star]
M_prime = [ep / ((ep ** 2) + (epp ** 2)) for ep, epp in zip(e_prime, e_double_prime)]
M_double_prime = [epp / ((ep ** 2) + (epp ** 2)) for ep, epp in zip(e_prime, e_double_prime)]
M_double_prime_max = max(M_double_prime)
M_double_prime_normalized = [Mp / M_double_prime_max for Mp in M_double_prime]

# Save results to Excel file
result_df = pd.DataFrame({
    'f': frequency,
    "Z' (Real Impedance)": Z_prime,
    "Z'' (Imaginary Impedance)": Z_double_prime,
    "Z''/Z''max": Z_double_prime_normalized,
    "e' (Dielectric Real)": e_prime,
    "e'' (Dielectric Imaginary)": e_double_prime,
    "e* (Complex Dielectric Value)": e_star,
    "tan delta (Dielectric Loss)": tan_delta,
    "M* (Complex Electric Modulus)": M_star,
    "M' (Modulus Real)": M_prime,
    "M'' (Modulus Imaginary)": M_double_prime,
    "M''/M''max": M_double_prime_normalized
})

result_df.to_excel('results.xlsx', index=False)

# Generate f vs Z' plot
fig, ax = plt.subplots()
ax.plot(frequency, Z_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("Z' (Real Impedance)")
ax.set_title("f vs Z'")
plt.savefig("f_vs_Z_prime.png")
plt.show()

# Generate f vs Z'' plot
fig, ax = plt.subplots()
ax.plot(frequency, Z_double_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("Z'' (Imaginary Impedance)")
ax.set_title("f vs Z''")
plt.savefig("f_vs_Z_double_prime.png")
plt.show()

# Generate f vs Z''/Z''max plot
fig, ax = plt.subplots()
ax.plot(frequency, Z_double_prime_normalized, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("Z''/Z''max")
ax.set_title("f vs Z''/Z''max")
plt.savefig("f_vs_Z_double_prime_normalized.png")
plt.show()

# Generate f vs e' plot
fig, ax = plt.subplots()
ax.plot(frequency, e_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("e' (Dielectric Real)")
ax.set_title("f vs e'")
plt.savefig("f_vs_e_prime.png")
plt.show()

# Generate f vs e'' plot
fig, ax = plt.subplots()
ax.plot(frequency, e_double_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("e'' (Dielectric Imaginary)")
ax.set_title("f vs e''")
plt.savefig("f_vs_e_double_prime.png")
plt.show()

# Generate f vs e* plot
fig, ax = plt.subplots()
ax.plot(frequency, e_star, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("e* (Complex Dielectric Value)")
ax.set_title("f vs e*")
plt.savefig("f_vs_e_star.png")
plt.show()

# Generate f vs tan delta plot
fig, ax = plt.subplots()
ax.plot(frequency, tan_delta, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("tan delta (Dielectric Loss)")
ax.set_title("f vs tan delta")
plt.savefig("f_vs_tan_delta.png")
plt.show()

# Generate f vs M* plot
fig, ax = plt.subplots()
ax.plot(frequency, M_star, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("M* (Complex Electric Modulus)")
ax.set_title("f vs M*")
plt.savefig("f_vs_M_star.png")
plt.show()

# Generate f vs M' plot
fig, ax = plt.subplots()
ax.plot(frequency, M_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("M' (Modulus Real)")
ax.set_title("f vs M'")
plt.savefig("f_vs_M_prime.png")
plt.show()

# Generate f vs M'' plot
fig, ax = plt.subplots()
ax.plot(frequency, M_double_prime, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("M'' (Modulus Imaginary)")
ax.set_title("f vs M''")
plt.savefig("f_vs_M_double_prime.png")
plt.show()

# Generate f vs M''/M''max plot
fig, ax = plt.subplots()
ax.plot(frequency, M_double_prime_normalized, marker='o')
ax.set_xlabel("Frequency (f)")
ax.set_ylabel("M''/M''max")
ax.set_title("f vs M''/M''max")
plt.savefig("f_vs_M_double_prime_normalized.png")
plt.show()
