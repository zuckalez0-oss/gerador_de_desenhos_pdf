import re
import math
import io
import ezdxf

def create_dxf_drawing(params: dict):
    """Gera um desenho DXF a partir de um dicionário de parâmetros já preparado."""
    try:
        doc = ezdxf.new('R2000')
        msp = doc.modelspace()
        
        styles = params.get('styles', {})
        doc.layers.new('CONTORNO', dxfattribs={'color': styles.get('contour_color', 7)})
        if params.get('holes'):
            doc.layers.new('FUROS', dxfattribs={'color': styles.get('holes_color', 7)})
        if params.get('text_lines'):
            doc.layers.new('TEXTO', dxfattribs={'color': styles.get('text_color', 7)})
        if styles.get('include_dims', False):
            doc.layers.new('COTAS', dxfattribs={'color': styles.get('text_color', 7)})
            doc.dimstyles.new('COTA-PADRAO', dxfattribs={'dimtxt': styles.get('char_height', 20)})
            dim_attribs = {'layer': 'COTAS', 'dimstyle': 'COTA-PADRAO'}

        shape_type = params.get('shape')
        
        shape_creators = {
            'rectangle': lambda m, p: m.add_lwpolyline([(0,0),(p.get('width', 0),0),(p.get('width', 0),p.get('height', 0)),(0,p.get('height', 0))], close=True),
            'circle': lambda m, p: m.add_circle(center=(p.get('diameter', 0)/2, p.get('diameter', 0)/2), radius=p.get('diameter', 0)/2),
            'right_triangle': lambda m, p: m.add_lwpolyline([(0,0),(p.get('rt_base', 0),0),(0,p.get('rt_height', 0))], close=True),
        }
        
        if shape_type in shape_creators:
            shape_creators[shape_type](msp, params).dxf.layer = 'CONTORNO'
        else:
            return None, f"Forma '{shape_type}' desconhecida."

        for hole in params.get('holes', []):
            msp.add_circle(center=(hole['x'], hole['y']), radius=hole['diameter']/2, dxfattribs={'layer': 'FUROS'})

        if params.get('text_lines'):
            start_point = styles.get('text_insert_point', (0, -20))
            char_height = styles.get('char_height', 5)
            line_spacing = char_height * 1.5
            for i, line in enumerate(params['text_lines']):
                y_pos = start_point[1] - (i * line_spacing)
                msp.add_text(line, dxfattribs={'layer': 'TEXTO', 'height': char_height, 'insert': (start_point[0], y_pos)})

        if styles.get('include_dims'):
            dim_distance = styles.get('dim_distance', 20)
            dims_creators = {
                'rectangle': lambda m, p: (
                    m.add_aligned_dim(p1=(0, p.get('height',0)), p2=(p.get('width',0), p.get('height',0)), distance=dim_distance, dxfattribs=dim_attribs).render(),
                    m.add_aligned_dim(p1=(0, 0), p2=(0, p.get('height',0)), distance=-dim_distance, dxfattribs=dim_attribs).render()
                ),
                 'circle': lambda m, p: m.add_diameter_dim(center=(p.get('diameter',0)/2, p.get('diameter',0)/2), radius=p.get('diameter',0)/2, angle=135, dxfattribs=dim_attribs).render(),
            }
            if shape_type in dims_creators:
                dims_creators[shape_type](msp, params)

        stream = io.StringIO()
        doc.write(stream)
        sanitized_filename = re.sub(r'[^\w.-]+', '_', str(params.get('part_name')))
        return stream.getvalue(), f"{sanitized_filename}.dxf"

    except Exception as e:
        print(f"Erro inesperado no desenho do DXF: {e}")
        return None, "Erro interno de desenho."

def prepare_and_validate_dxf_data(raw_data: dict):
    """Prepara e valida os dados para a geração do DXF."""
    params = raw_data.copy()

    # Mapeamento de nomes de colunas do PyQt para os nomes esperados pela engine
    params['part_name'] = params.get('nome_arquivo')
    params['shape'] = params.get('forma')
    params['width'] = params.get('largura')
    params['height'] = params.get('altura')
    params['diameter'] = params.get('diametro')
    params['part_quantity'] = params.get('qtd')
    params['material_thickness'] = params.get('espessura')
    
    # --- MUDANÇA CRUCIAL: DESLIGA AS COTAS E O TEXTO PARA O DXF ---
    # O DXF gerado será "limpo", contendo apenas a geometria para a máquina.
    params['include_dims'] = False 
    params['include_text_info'] = False
    
    if not params.get('part_name') or not params.get('shape'):
        return None, f"Dados insuficientes: 'part_name' ou 'shape' ausentes."

    def to_float(value, default=0.0):
        if value is None or value == '': return default
        try: return float(str(value).replace(',', '.'))
        except (ValueError, TypeError): return default

    numeric_keys = ['width', 'height', 'diameter', 'material_thickness', 'part_quantity', 'rt_base', 'rt_height']
    for key in numeric_keys:
        params[key] = to_float(params.get(key))
        
    params['holes'] = []
    if 'furos' in params and isinstance(params['furos'], list):
        for furo_pyqt in params['furos']:
            params['holes'].append({
                'diameter': to_float(furo_pyqt.get('diam')),
                'x': to_float(furo_pyqt.get('x')),
                'y': to_float(furo_pyqt.get('y'))
            })

    max_dim = max(params.get('width', 0), params.get('height', 0), params.get('diameter', 0), 100)
    char_height = min(max(max_dim / 25, 5), 35)
    
    params['styles'] = {
        'contour_color': 1, # Vermelho
        'holes_color': 3,   # Verde
        'text_color': 7,    # Branco/Preto
        'include_dims': params.get('include_dims'), 
        'char_height': char_height,
        'dim_distance': max(15, char_height * 3), 
        'text_insert_point': (char_height, -char_height * 2),
    }
    
    # Como 'include_text_info' é False, este bloco não será executado,
    # e 'text_lines' não será adicionado aos parâmetros.
    if params.get('include_text_info'):
        params['text_lines'] = [
            f"{str(params.get('part_name', '')).upper()}",
            f"Esp: {params.get('material_thickness', 0):.2f}mm (Qtd: {int(params.get('part_quantity', 1))})",
        ]
            
    return params, None