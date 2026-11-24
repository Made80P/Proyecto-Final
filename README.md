# Proyecto-Final
Madeleine palacios Baltan
Isabella Moreno Gomez
# Descripcion
Aplicación de escritorio para la gestión y venta de productos de pastelería artesanal Duleittes, desarrollada en Python con Tkinter. Incluye catálogo, búsqueda y filtros, registro e inicio de sesión de clientes, carrito de compras y módulo administrativo.

# Problema 
la pastelería Duleittes actualmente presenta dificultades para organizar y centralizar la información de sus productos, debido a que la interacción con los clientes y la promoción se realizan principalmente mediante redes sociales y canales digitales. Esto puede dificultar la presentación del catálogo, la consulta de productos y precios, así como la gestión de pedidos. Además, a medida que el emprendimiento busca crecer, surge la necesidad de contar con un espacio digital propio que facilite estos procesos y mejore la interacción con los clientes.
<img width="938" height="703" alt="image" src="https://github.com/user-attachments/assets/a085f774-4597-4371-a6f4-ccd406a09ef5" />

# Usuario
El usuario objetivo de Duleittes son personas jóvenes-adultas y familias que valoran los productos artesanales, el buen sabor, la calidad y una presentación atractiva. Buscan opciones diferentes a lo tradicional para compartir, celebrar o disfrutar momentos especiales, y prefieren la facilidad de conocer y adquirir los productos mediante medios digitales

# Objetivos
Como equipo, decidimos que nuestros objetivos principales son: organizar y presentar de manera clara el catálogo de productos, facilitando que los clientes conozcan sus características, precios e ingredientes. También buscamos facilitar la compra, con opciones de tamaño, adicionales, carrito y cálculo automático del precio. Finalmente, queremos mejorar la administración mediante el control del inventario, alertas de stock y reportes de ventas.

# Requerimientos Funcionales – Módulo de Cliente
Un catálogo interactivo organizado por categorías, con imágenes, precios, ingredientes y stock; un buscador en tiempo real; la personalización de productos como tortas y galletas con recálculo automático del precio; un carrito de compras con cupones; la generación de comprobantes digitales; y diferentes secciones informativas como términos, políticas de privacidad, preguntas frecuentes y productos más vendidos.

# Requerimientos Funcionales – Módulo Administrativo
Tener autenticación protegida, gestión del inventario mediante un sistema CRUD en Excel para registrar, actualizar y eliminar productos, alertas visuales de stock crítico, gestión automática de las imágenes y reportes diarios de ventas que permiten consultar los ingresos, número de pedidos y producto más vendido.

# Requerimientos Técnicos y No Funcionales
Contar con archivos de Excel y archivos de texto para almacenar la información. También contamos con validación de entradas y manejo de errores, por ejemplo, al guardar el catálogo. Finalmente, el sistema sincroniza automáticamente las vistas de la tienda y el inventario cuando se realizan cambios desde el panel administrativo.

# Funcionamiento del backend
El backend es la parte encargada de procesar y controlar la información de la aplicación. En nuestro proyecto gestiona los productos, clientes, pedidos, inventario, carrito de compras y autenticación.
Utilizamos archivos de Excel y texto para almacenar y consultar la información. El backend permite buscar y filtrar productos, calcular precios según las opciones seleccionadas, aplicar descuentos, controlar el stock y generar comprobantes de compra.
Además, desde el módulo administrativo se pueden registrar, actualizar y eliminar productos, y los cambios se sincronizan con la tienda. De esta manera, el backend conecta la interfaz con los datos y permite que las diferentes funciones del sistema trabajen de forma organizada y actualizada.

# Funcionamiento del frontend
Desarrollada con Tkinter. Es la parte con la que interactúa directamente el usuario y permite navegar por la aplicación de manera visual.
Desde esta interfaz, el cliente puede explorar el catálogo por categorías, visualizar las imágenes y características de los productos, utilizar el buscador, personalizar sus pedidos y agregarlos al carrito. También puede acceder a las opciones de registro, inicio de sesión, cupones y comprobante de compra.
Por otro lado, el administrador cuenta con una interfaz propia desde la cual puede acceder a las funciones de gestión del catálogo e inventario.
En resumen, el frontend se encarga de mostrar la información y recibir las acciones del usuario, haciendo que todas las funciones de Duleittes sean accesibles mediante una interfaz gráfica.

# Validación del Usuario
Realizamos entrevistas a diferentes usuarios para evaluar la facilidad de uso y la funcionalidad de la aplicación. Los participantes lograron completar las tareas sin ayuda y destacaron que la navegación, los botones y la información son claros y fáciles de entender.
También se validaron funciones importantes como el registro e inicio de sesión, catálogo de productos, consulta de ingredientes, carrito de compras y proceso de pedido. Los usuarios consideraron que estas funciones facilitan y agilizan la compra.
Como resultado, la mayoría indicó que utilizaría la aplicación en una situación real y que podría ser una solución útil para emprendimientos o panaderías, confirmando que la propuesta responde a las necesidades identificadas.
<img width="296" height="223" alt="image" src="https://github.com/user-attachments/assets/efedc58f-ee42-4703-9d60-361aafd1476b" />
<img width="293" height="219" alt="image" src="https://github.com/user-attachments/assets/73a4f240-b078-44b0-b4c3-09585b6b377b" />
