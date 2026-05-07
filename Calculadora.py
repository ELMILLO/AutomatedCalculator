import xlwings as xw

XL_CENTER = -4108
XL_CONTINUOUS = 1
XL_BORDER_WEIGHT = 2

def _aplicar_estilo(rng, color=None, es_titulo=False, es_cabecera=False):
    rng.api.HorizontalAlignment = XL_CENTER
    rng.api.VerticalAlignment = XL_CENTER
    if color: rng.color = color
    if es_titulo:
        rng.api.Font.Size = 16
        rng.api.Font.Bold = True
        rng.api.Font.Color = 0x000000 
    if es_cabecera: rng.api.Font.Bold = True
    
    if not es_titulo:
        for i in [7, 8, 9, 10, 11]:
            border = rng.api.Borders(i)
            border.LineStyle = XL_CONTINUOUS
            border.Weight = XL_BORDER_WEIGHT

def obtener_despiece(mueble, mat_est, mat_pte, ancho, alto, prof, sobrada, cant_entrep):
    piezas = []
    desc = 30 
    # Regla de pieza sobrada suma 15mm al alto base si aplica (para otros muebles)
    alto_p = alto + 15 if str(sobrada).strip().lower() == "sí" else alto
    ancho_val = ancho if ancho else 0
    num_puertas = 2 if ancho_val > 500 else 1
    ancho_puerta = (ancho_val - 4) / 2 if num_puertas == 2 else (ancho_val - 4)

    def add(tipo_mat, largo, anch, cant, ref, l1, l2, w1, w2):
        mat = mat_est if tipo_mat == "est" else mat_pte
        piezas.append([mat, round(largo, 2), round(anch, 2), cant, ref, l1, l2, w1, w2, mat, mueble])

    if mueble == "Cajonera 2 cajones":
        add("est", alto, prof, 2, "costados", 1, 0, 0, 0)
        add("est", prof, ancho - desc, 1, "piso", 0, 0, 1, 0)
        add("est", 100, ancho - desc, 2, "fondo", 0, 0, 1, 1)
        add("est", 100, ancho - desc, 2, "techo", 0, 0, 1, 1)
        add("est", 420, ancho - 86, 2, "fondo cajón", 0, 0, 0, 0)
        add("est", 450, 160, 4, "laterales cajón", 1, 0, 0, 0)
        add("est", ancho - 86, 160, 4, "frontal cajón", 1, 0, 0, 0)
        frente_largo = (alto - 70) / 2
        add("pte", frente_largo, ancho - 5, 2, "frentes cajón", 1, 1, 1, 1)

    elif mueble == "Cajonera 3 cajones":
        add("est", alto, prof, 2, "costados", 1, 0, 0, 0)
        add("est", prof, ancho - desc, 1, "piso", 0, 0, 1, 0)
        add("est", 100, ancho - desc, 2, "fondo", 0, 0, 1, 1)
        add("est", 100, ancho - desc, 2, "techo", 0, 0, 1, 1)
        add("est", 420, ancho - 86, 2, "fondo cajón", 0, 0, 0, 0)
        add("est", 450, 160, 4, "laterales cajón", 1, 0, 0, 0)
        add("est", ancho - 86, 160, 4, "frontal cajón", 1, 0, 0, 0)
        add("est", 420, ancho - 86, 1, "fondo cajón int", 0, 0, 0, 0)
        add("est", 450, 80, 2, "laterales cajón int", 1, 0, 0, 0)
        add("est", ancho - 86, 80, 2, "frontal cajón int", 1, 0, 0, 0)
        frente_largo = (alto - 70) / 2
        frente_int_largo = alto - 620
        add("pte", frente_largo, ancho - 4, 2, "frentes cajón", 1, 1, 1, 1)
        add("pte", frente_int_largo, ancho - 34, 1, "frente cajón int", 1, 1, 1, 1)

    elif mueble == "Gabinete":
        add("est", alto, prof, 2, "costados", 1, 0, 0, 0)
        add("est", prof, ancho - desc, 2, "techo y piso", 0, 0, 1, 0)
        add("est", alto - desc, ancho - desc, 1, "fondo", 0, 0, 0, 0)
        try: ce = int(cant_entrep)
        except: ce = 0
        if ce > 0:
            add("est", prof - 15, ancho - desc, ce, "entrepaño", 0, 0, 0, 1)
            
        # Corrección exacta de tu fórmula: Las puertas del gabinete siempre son Alto + 15
        largo_puerta_gabinete = alto + 15
        add("pte", largo_puerta_gabinete, ancho_puerta, num_puertas, "puertas", 1, 1, 1, 1)

    elif mueble == "Tarja":
        add("est", alto, prof, 2, "costados", 1, 0, 0, 0)
        add("est", prof, ancho - desc, 1, "piso", 0, 0, 1, 0)
        add("est", 100, ancho - desc, 2, "manguetes sup.", 0, 0, 1, 1)
        largo_puerta = alto_p - 40
        add("pte", largo_puerta, ancho_puerta, num_puertas, "puertas", 1, 1, 1, 1)

    elif mueble == "Corte Personalizado":
        add("est", alto, ancho, 1, mueble, 1, 1, 1, 1)
        
    return piezas

@xw.sub
def calcular_avance_viernes():
    try:
        wb = xw.Book.caller()
        sh_ui = wb.sheets['Calculadora']
        mueble = str(sh_ui.range('C4').value).strip()
        mat_est = str(sh_ui.range('C6').value).strip()
        mat_pte = str(sh_ui.range('C7').value).strip()
        ancho = sh_ui.range('C8').value
        prof = sh_ui.range('C9').value
        alto = sh_ui.range('C10').value
        cant_entrep = sh_ui.range('C11').value
        sobrada = sh_ui.range('C12').value
        proyecto = str(sh_ui.range('C14').value).strip()

        if mueble == "Corte Personalizado":
            if prof is not None and str(prof).strip() != "":
                sh_ui.range('E15').value = "⚠️ Error: Es 2D. Borra la Profundidad."
                sh_ui.range('E15').color = (255, 199, 206)
                return
        elif not proyecto or proyecto == "None" or not ancho or not alto or not prof:
            sh_ui.range('E15').value = "⚠️ Faltan Medidas o Proyecto"
            sh_ui.range('E15').color = (255, 199, 206)
            return

        if proyecto in [s.name for s in wb.sheets]: 
            sh_out = wb.sheets[proyecto]
        else:
            sh_out = wb.sheets.add(proyecto, after=sh_ui)
            sh_out.range('A1:K1').api.Merge()
            sh_out.range('A1:K1').value = proyecto
            _aplicar_estilo(sh_out.range('A1:K1'), es_titulo=True)
            headers = ["Material", "Largo", "Ancho", "Cantidad", "Referencia de corte", "Lado largo1", "Lado Largo 2", "Lado ancho 1", "Lado ancho 2", "Color enchape", "Nombre del mueble"]
            sh_out.range('A2:K2').value = headers
            _aplicar_estilo(sh_out.range('A2:K2'), color=(255, 255, 0), es_cabecera=True)

        lista_piezas = obtener_despiece(mueble, mat_est, mat_pte, ancho, alto, prof, sobrada, cant_entrep)
        if not lista_piezas:
            return

        last_row = sh_out.range('A' + str(sh_out.cells.last_cell.row)).end('up').row
        start = 3 if last_row < 2 else last_row + 1
        sh_out.range(f'A{start}').value = lista_piezas
        end = start + len(lista_piezas) - 1
        _aplicar_estilo(sh_out.range(f'A{start}:K{end}'))
        sh_out.range('A:K').autofit()
        
        wb.app.display_alerts = False
        m_rng = sh_out.range(f'K{start}:K{end}')
        m_rng.api.Merge()
        m_rng.api.VerticalAlignment = m_rng.api.HorizontalAlignment = XL_CENTER
        wb.app.display_alerts = True

        sh_ui.range('E15').value, sh_ui.range('E15').color = f"✅ {mueble} agregado!", (200, 255, 200)
    except Exception as e:
        sh_ui.range('E15').value, sh_ui.range('E15').color = f"Error: {str(e)}", (255, 199, 206)
