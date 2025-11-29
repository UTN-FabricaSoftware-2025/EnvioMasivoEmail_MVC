Feature: Incrustación de Banner en HTML de Correo
  Como servicio de envío de correos
  Quiero procesar el cuerpo HTML para incrustar una imagen local si existe
  Para asegurar que el destinatario vea el banner correctamente o un texto alternativo

  Background: Configuración inicial
    # Se inicializa un BodyBuilder vacío antes de cada prueba
    Given que inicializo un nuevo objeto BodyBuilder
    And defino la ruta de la imagen como "wwwroot/Plantillas/img2.jpg"

  Scenario Outline: Validación de lógica de reemplazo e incrustación de imagen
    # Simulamos la existencia del archivo físico para probar los caminos de código
    Given que el archivo de imagen en la ruta definida <estado_archivo>
    And el cuerpo HTML de entrada es "<html_entrada>"
    When ejecuto la función interna IncrustarBanner
    Then el HTML resultante debe contener el texto "<texto_esperado>"
    And el objeto Builder debe tener <adjuntos_count> recursos vinculados (LinkedResources)

    Examples: Casos de prueba (Escenarios A, B y C)
      | estado_archivo | html_entrada                     | texto_esperado           | adjuntos_count | Notas                                                                                                |
      | existe         | <div>{{img2}}</div>              | <img src="cid:           | 1              | ESCENARIO A (Camino Feliz): Imagen existe, se reemplaza el token por CID y se adjunta al Builder.    |
      | no existe      | <div>{{img2}}</div>              | IMG no encontrada        | 0              | ESCENARIO B (Imagen Falta): No existe archivo, se reemplaza token por texto error. Nada se adjunta.  |
      | existe         | <h1>Hola Mundo</h1>              | <h1>Hola Mundo</h1>      | 1              | ESCENARIO C (Sin Token): Imagen existe y se adjunta (desperdicio), pero HTML no cambia al no tener {{img2}}. |