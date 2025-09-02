import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# --- Parameters ---
Z0 = 50  # Characteristic impedance in Ohms
filename = "C:\Users\verma\Downloads\NEW_COIL.s1p"  # Replace with your file path
skip_header = 12  # Adjust depending on your file
output_excel = "calculated_impedance.xlsx"

# --- Load and parse data ---
data = np.loadtxt(filename, skiprows=skip_header)
frequency = data[:, 0]
mag_dB = data[:, 1]
angle_deg = data[:, 2]

# --- Convert to complex reflection coefficient ---
mag_lin = 10**(mag_dB / 20)
angle_rad = np.radians(angle_deg)
gamma = mag_lin * np.exp(1j * angle_rad)

# --- Calculate impedance ---
Z = Z0 * (1 + gamma) / (1 - gamma)
real_Z = np.real(Z)
imag_Z = np.imag(Z)

# --- Export to Excel ---
df = pd.DataFrame({
    'FREQUENCY(Hz)': frequency,
    'R': real_Z,
    'L': imag_Z
})

df.to_excel(output_excel, index=False, engine='openpyxl')
print(f"Impedance data exported to '{output_excel}'")

# --- Plot impedance ---
plt.figure(figsize=(10, 5))
plt.plot(frequency / 1e6, real_Z, label='Real(Z)')
plt.plot(frequency / 1e6, imag_Z, label='Imag(Z)')
plt.xlabel("Frequency [MHz]")
plt.ylabel("Impedance [Ohm]")
plt.title("Load Impedance from Reflection Coefficient")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
