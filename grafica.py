import matplotlib.pyplot as plt
import numpy as np

x = [0.0, 0.68, 2.04, 2.72, 3.4, 4.08, 4.76, 5.44, 6.12, 6.8, 7.48, 8.16, 8.4, 8.4]
y1 = [0.96, 1.0, 1.0, 1.0, 1.05, 0.9, 0.8, 0.67, 0.65, 0.45, 0.26, 0.1, 0.0, 0.0]
y2 = [0.0, 0.23359000000000002, 0.3668376, 0.4092468, 0.2880092, 0.2691408, 0.3489612, 0.45309360000000004, 0.4221852, 0.3347312, 0.32802240000000005, 0.2666848, 0.0, 0.0]

ruta = r"C:\Users\SHI-PC34.SHI-PC34\Desktop\BBBBBBBBBBBB.png"
fig, ax1 = plt.subplots(figsize=(12,7))
ax2 = ax1.twinx()
fig.text(0.13, 0.01, 'MARGEN DERECHA', size = 12)
ax1.plot(x,y1)
ax2.bar(x,y2, color = 'orange', edgecolor = 'brown', linewidth = 1.5, alpha = 0.6, label = 'Velocidad', width = 0.1)
ax1.fill_between(x, y1, color='skyblue')
ax2.invert_yaxis() 
ax1.invert_yaxis()
ax1.set_xticks(x)
ax1.set_xticklabels(x, rotation=45)
ax1.tick_params(axis = 'x', labelsize = 11)
ax1.tick_params(axis = 'y', labelsize = 11)
ax2.tick_params(axis = 'y', labelsize = 11)

textura_1 = plt.imread(r'C:\Users\SHI-PC34.SHI-PC34\Desktop\Martín Anaya\01 Aforos Líquidos\02 VSC\01 Input\textura-puntos.jpg')
ax1.imshow(textura_1, extent=[ax1.get_xlim()[0], ax1.get_xlim()[1], ax1.get_ylim()[0], ax1.get_ylim()[1]], aspect='auto', alpha=0.5, cmap='gray', interpolation='bilinear')

ax1.set_title('Perfil de profundidad y velocidad', fontweight='bold', fontsize=16, pad = 20)
ax1.set_xlabel('Ancho (m)', fontweight='bold', fontsize=14, labelpad = 20)
ax1.set_ylabel('Profundidad (m)', fontweight='bold', fontsize=14, labelpad = 20)
ax2.set_ylabel('Velocidad del flujo (m/s)', fontweight='bold', fontsize=14, labelpad = 20)

plt.legend()
plt.show()
plt.savefig(ruta, dpi = 600)



