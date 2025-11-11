import streamlit as st
import json
import re
from collections import defaultdict
import pandas as pd
from datetime import datetime
import io
import xlsxwriter

def parse_assessment_file(file_content):
    dimensions = []
    lines = file_content.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            i += 1
            continue
        
        if re.match(r'^\d+\.', line) and ':' in line:
            parts = line.split(':', 1)
            if len(parts) == 2:
                after_colon = parts[1].strip()
                if len(after_colon) > 10 and (re.search(r'[A-Z]', after_colon) or '-' in after_colon):
                    dimension_name = line
                    questions = []
                    answers = {}
                    i += 1
                    
                    while i < len(lines):
                        line = lines[i].strip()
                        
                        if line.startswith('Answers:'):
                            i += 1
                            answer_list = []
                            while i < len(lines):
                                line = lines[i].strip()
                                if line.startswith('*'):
                                    answer_text = line[1:].strip()
                                    if ':' in answer_text:
                                        parts = answer_text.split(':', 1)
                                        if len(parts) >= 2:
                                            key = parts[0].strip()
                                            value = parts[1].strip()
                                            answers[key] = value
                                    else:
                                        answer_list.append(answer_text)
                                    i += 1
                                elif not line:
                                    i += 1
                                    break
                                elif re.match(r'^\d+\.', line) and ':' in line:
                                    test_parts = line.split(':', 1)
                                    if len(test_parts) == 2:
                                        test_after = test_parts[1].strip()
                                        if len(test_after) > 10 and (re.search(r'[A-Z]', test_after) or '-' in test_after):
                                            break
                                    i += 1
                                else:
                                    i += 1
                            
                            if answer_list and not answers:
                                for idx, value in enumerate(answer_list):
                                    answers[chr(65 + idx)] = value
                            break
                        
                        if re.match(r'^\d+\.', line) and ':' in line:
                            test_parts = line.split(':', 1)
                            if len(test_parts) == 2:
                                test_after = test_parts[1].strip()
                                if len(test_after) > 10 and (re.search(r'[A-Z]', test_after) or '-' in test_after):
                                    break
                        
                        if re.match(r'^Question\s+\d+:', line, re.IGNORECASE):
                            question_text = re.sub(r'^Question\s+\d+:\s*', '', line, flags=re.IGNORECASE)
                            options = []
                            i += 1
                            
                            while i < len(lines):
                                line = lines[i].strip()
                                if re.match(r'^[a-gA-G]\)', line):
                                    option = re.sub(r'^[a-gA-G]\)\s*', '', line)
                                    options.append(option)
                                    i += 1
                                else:
                                    break
                            
                            if options:
                                questions.append({
                                    'text': question_text,
                                    'options': options
                                })
                        elif re.match(r'^\d+\.', line):
                            question_text = re.sub(r'^\d+\.\s*', '', line)
                            options = []
                            i += 1
                            
                            while i < len(lines):
                                line = lines[i].strip()
                                if re.match(r'^[a-gA-G]\)', line):
                                    option = re.sub(r'^[a-gA-G]\)\s*', '', line)
                                    options.append(option)
                                    i += 1
                                else:
                                    break
                            
                            if options:
                                questions.append({
                                    'text': question_text,
                                    'options': options
                                })
                        else:
                            i += 1
                    
                    if questions:
                        dimensions.append({
                            'name': dimension_name,
                            'questions': questions,
                            'answers': answers
                        })
                    continue
        
        i += 1
    
    return dimensions

def calculate_results(responses, dimensions):
    results = {}
    
    for dim_idx, dimension in enumerate(dimensions):
        if dim_idx not in responses:
            continue
        
        type_counts = defaultdict(int)
        answers_map = dimension['answers']
        
        for q_idx, answer in responses[dim_idx].items():
            if answer:
                answer_upper = answer.upper()
                if answer_upper in answers_map:
                    answer_type = answers_map[answer_upper]
                    type_counts[answer_type] += 1
        
        if type_counts:
            dominant_type = max(type_counts.items(), key=lambda x: x[1])
            total = sum(type_counts.values())
            percentage = (dominant_type[1] / total * 100) if total > 0 else 0
            
            results[dimension['name']] = {
                'dominant_type': dominant_type[0],
                'percentage': percentage,
                'all_scores': dict(type_counts),
                'total_questions': total
            }
    
    return results

def generate_personal_profile(cognitive_results, conative_results, semantic_results, emotional_results):
    """Generate comprehensive personal profile based on assessment results"""
    
    profile_sections = []
    
    # Extract key traits
    cog_traits = {dim.split(':')[0].strip(): result['dominant_type'] 
                  for dim, result in cognitive_results.items()}
    con_traits = {dim.split(':')[0].strip(): result['dominant_type'] 
                  for dim, result in conative_results.items()}
    sem_traits = {dim.split(':')[0].strip(): result['dominant_type'] 
                  for dim, result in semantic_results.items()}
    emo_traits = {dim.split(':')[0].strip(): result['dominant_type'] 
                  for dim, result in emotional_results.items()}
    
    # GENERAL PERSONALITY
    personality_lines = ["GENERAL PERSONALITY:", ""]
    
    # Cognitive style
    rep_type = cog_traits.get('1. Representation', 'Visual')
    personality_lines.append(f"Your primary cognitive processing style is {rep_type}, meaning you best understand and retain information through {rep_type.lower()} channels.")
    
    # Learning approach
    epistemo = cog_traits.get('2. Epistemological', 'Sensor')
    if 'Sensor' in epistemo:
        personality_lines.append("You are detail-oriented and prefer concrete, tangible information with practical applications.")
    else:
        personality_lines.append("You are conceptual and prefer abstract thinking, seeing patterns and possibilities.")
    
    # Decision making
    scale = cog_traits.get('3. Scale', 'Global')
    if 'Global' in scale or 'Inductive' in scale:
        personality_lines.append("You tend to see the big picture first before diving into details.")
    else:
        personality_lines.append("You prefer starting with specifics and building up to comprehensive understanding.")
    
    # Emotional style
    surgency = emo_traits.get('3. Exuberance', 'Surgency')
    if 'Surgency' in surgency:
        personality_lines.append("You have a bold, confident presence and aren't afraid to speak up.")
    else:
        personality_lines.append("You have a reserved, thoughtful demeanor and prefer listening before speaking.")
    
    profile_sections.append("\n".join(personality_lines))
    
    # SUITABLE CAREER PATHS
    career_lines = ["", "", "SUITABLE CAREER PATHS:", ""]
    
    work_style = emo_traits.get('12. Work Style', 'Independent')
    dominance = emo_traits.get('11. Dominance', 'Achievement')
    focus = cog_traits.get('9. Focus', 'Screening')
    
    if 'Independent' in work_style:
        career_lines.append("Best suited for roles with autonomy: Research, Consulting, Freelancing, Entrepreneurship")
    elif 'Manager' in work_style or 'Leader' in work_style:
        career_lines.append("Natural leadership abilities: Management, Executive roles, Team Leadership, Project Management")
    elif 'Team player' in work_style:
        career_lines.append("Collaborative environments: Team-based roles, Partnership positions, Cooperative projects")
    
    if 'Achievement' in dominance:
        career_lines.append("Goal-oriented roles where you can master skills and deliver high-quality outcomes")
    elif 'Power' in dominance:
        career_lines.append("Leadership and influence positions where you can drive initiatives and make decisions")
    elif 'Affiliation' in dominance:
        career_lines.append("People-focused roles: HR, Counseling, Team Building, Community Management")
    
    # Add specific field recommendations
    career_lines.append("")
    career_lines.append("Recommended Fields:")
    
    if 'Visual' in rep_type:
        career_lines.append("- Design, Architecture, Data Visualization, Photography, UI/UX")
    if 'Auditory' in rep_type:
        career_lines.append("- Music, Voice Acting, Audio Engineering, Teaching, Public Speaking")
    if 'Kinesthetic' in rep_type:
        career_lines.append("- Athletics, Physical Therapy, Surgery, Hands-on Crafts, Dance")
    
    profile_sections.append("\n".join(career_lines))
    
    # SOCIAL INTERACTIONS
    social_lines = ["", "", "SOCIAL INTERACTIONS:", ""]
    
    attention = emo_traits.get('6. Attention', 'Self')
    rejuvenation = emo_traits.get('8. Rejuvenation', 'Introvert')
    presentation = emo_traits.get('10. Societal Presentation', 'Genuine')
    
    if 'Introvert' in rejuvenation:
        social_lines.append("You recharge through solitude and prefer smaller, intimate gatherings over large social events.")
    else:
        social_lines.append("You gain energy from social interactions and thrive in group settings.")
    
    if 'Self' in attention:
        social_lines.append("In conversations, you tend to focus on your own perspective and experiences.")
    else:
        social_lines.append("You are naturally attuned to others' needs and perspectives in social situations.")
    
    if 'Genuine' in presentation or 'Artlessly' in presentation:
        social_lines.append("You value authenticity and prefer straightforward, honest communication.")
    else:
        social_lines.append("You are diplomatically skilled and adapt your approach based on social context.")
    
    profile_sections.append("\n".join(social_lines))
    
    # RELATIONSHIPS
    relationship_lines = ["", "", "RELATIONSHIPS:", ""]
    
    self_exp = sem_traits.get('1. Self-Experience', 'Mind')
    emotional_contain = emo_traits.get('7. Emotional Containment', 'Contain')
    movie_pos = emo_traits.get('2. Movie Position', 'Associated')
    
    if 'Emotions' in self_exp:
        relationship_lines.append("You connect with others primarily through emotional bonds and shared feelings.")
    elif 'Mind' in self_exp:
        relationship_lines.append("You build relationships through intellectual connection and meaningful conversations.")
    elif 'Body' in self_exp:
        relationship_lines.append("Physical presence and shared activities form the foundation of your relationships.")
    
    if 'Contain' in emotional_contain:
        relationship_lines.append("You tend to keep emotions private, sharing deep feelings only with those closest to you.")
    elif 'Spread' in emotional_contain:
        relationship_lines.append("You openly share your emotions and appreciate when others do the same.")
    
    if 'Associated' in movie_pos:
        relationship_lines.append("You experience relationships fully, being emotionally present and engaged in the moment.")
    else:
        relationship_lines.append("You maintain some emotional distance, which helps you stay objective in relationships.")
    
    profile_sections.append("\n".join(relationship_lines))
    
    # COMMUNICATION STYLE
    comm_lines = ["", "", "COMMUNICATION STYLE:", ""]
    
    communication = cog_traits.get('11. Communication', 'Verbal')
    stress_coping = emo_traits.get('4. Stress Coping', 'Assertive')
    
    if 'Verbal' in communication or 'Digital' in communication:
        comm_lines.append("You communicate best through words, whether spoken or written.")
    else:
        comm_lines.append("You communicate through non-verbal cues, tone, and body language as much as words.")
    
    if 'Passive' in stress_coping:
        comm_lines.append("You tend to avoid confrontation and may need time to process before responding.")
    elif 'Assertive' in stress_coping:
        comm_lines.append("You express your needs clearly while respecting others' perspectives.")
    elif 'Aggressive' in stress_coping:
        comm_lines.append("Under stress, you may become forceful in expressing your viewpoint.")
    
    profile_sections.append("\n".join(comm_lines))
    
    # DECISION MAKING & VALUES
    decision_lines = ["", "", "DECISION MAKING & VALUES:", ""]
    
    authority = emo_traits.get('5. Authority Source', 'Internal')
    self_instruction = sem_traits.get('2. Self -Instruction', 'Neutral')
    quality_life = sem_traits.get('12. Quality of Life', 'Be')
    
    if 'Internal' in authority:
        decision_lines.append("You trust your own judgment and make decisions based on internal values and beliefs.")
    else:
        decision_lines.append("You value external input and seek advice from others when making important decisions.")
    
    if 'Strong Will' in self_instruction:
        decision_lines.append("You prefer autonomy and resist following instructions that don't align with your approach.")
    elif 'Compliant' in self_instruction:
        decision_lines.append("You respect structure and are comfortable following established guidelines.")
    
    if 'Be' in quality_life:
        decision_lines.append("You prioritize personal growth and self-awareness over external achievements.")
    elif 'Do' in quality_life:
        decision_lines.append("You find fulfillment in accomplishments and tangible results.")
    elif 'Have' in quality_life:
        decision_lines.append("You value material success and the accumulation of resources.")
    
    profile_sections.append("\n".join(decision_lines))
    
    # STRENGTHS & GROWTH AREAS
    strengths_lines = ["", "", "KEY STRENGTHS:", ""]
    
    confidence = sem_traits.get('3. Self Confidence', 'High')
    esteem = sem_traits.get('4. Self Esteem', 'Unconditional')
    ego = sem_traits.get('7. Ego Strength', 'Strong')
    
    if 'High' in confidence:
        strengths_lines.append("- Strong self-confidence and belief in your abilities")
    if 'Unconditional' in esteem:
        strengths_lines.append("- Stable self-worth not dependent on external validation")
    if 'Strong' in ego:
        strengths_lines.append("- Resilient in face of criticism and setbacks")
    
    change = emo_traits.get('13. Change Adapter', 'Medium')
    if 'Early' in change:
        strengths_lines.append("- Quick to adapt and embrace innovation")
    
    responsibility = sem_traits.get('6. Responsibility', 'Responsible')
    if 'Responsible' in responsibility and 'Over' not in responsibility and 'Under' not in responsibility:
        strengths_lines.append("- Balanced sense of accountability")
    
    profile_sections.append("\n".join(strengths_lines))
    
    # TIME ORIENTATION
    time_lines = ["", "", "TIME ORIENTATION:", ""]
    
    time_zone = sem_traits.get('10. Time Zones', 'Present')
    time_exp = sem_traits.get('11. Time Experience', 'In Time')
    
    if 'Past' in time_zone:
        time_lines.append("You often reflect on past experiences and learn from history.")
    elif 'Present' in time_zone:
        time_lines.append("You focus on the here and now, making the most of current opportunities.")
    elif 'Future' in time_zone:
        time_lines.append("You are forward-thinking, constantly planning and preparing for what's ahead.")
    
    if 'Sequential' in time_exp:
        time_lines.append("You approach tasks in an organized, step-by-step manner.")
    elif 'Random' in time_exp:
        time_lines.append("You think holistically and may jump between ideas non-linearly.")
    
    profile_sections.append("\n".join(time_lines))
    
    # STRESS MANAGEMENT
    stress_lines = ["", "", "STRESS MANAGEMENT:", ""]
    
    response = emo_traits.get('9. Somatic Response', 'Reflective')
    persistence = emo_traits.get('15. Persistence', 'Patient')
    attitude = emo_traits.get('14. Attitude', 'Serious')
    
    if 'Reflective' in response:
        stress_lines.append("You handle stress by pausing to think through situations before acting.")
    else:
        stress_lines.append("You manage stress by taking immediate action to address problems.")
    
    if 'Patient' in persistence:
        stress_lines.append("You have the patience to work through long-term challenges steadily.")
    else:
        stress_lines.append("You prefer quick results and may become restless with slow progress.")
    
    if 'Playful' in attitude:
        stress_lines.append("You use humor and lightheartedness to diffuse tension.")
    else:
        stress_lines.append("You maintain focus and determination when facing difficulties.")
    
    profile_sections.append("\n".join(stress_lines))
    
    return "\n".join(profile_sections)

def create_excel_report(client_name, cognitive_responses, cognitive_dimensions, 
                       conative_responses, conative_dimensions,
                       semantic_responses, semantic_dimensions,
                       emotional_responses, emotional_dimensions):
    buffer = io.BytesIO()
    
    # Calculate results for all sections
    cognitive_results = calculate_results(cognitive_responses, cognitive_dimensions)
    conative_results = calculate_results(conative_responses, conative_dimensions)
    semantic_results = calculate_results(semantic_responses, semantic_dimensions)
    emotional_results = calculate_results(emotional_responses, emotional_dimensions)
    
    # Create workbook
    workbook = xlsxwriter.Workbook(buffer, {'constant_memory': False})
    
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4CAF50',
        'font_color': 'white',
        'border': 1
    })
    
    text_wrap_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top'
    })
    
    # 1. SUMMARY SHEET
    worksheet = workbook.add_worksheet('Summary')
    worksheet.set_column('A:A', 60)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 18)
    
    headers = ['Dimension', 'Dominant Type', 'Strength (%)', 'Total Questions']
    for col, header in enumerate(headers):
        worksheet.write(1, col, header, header_format)
    
    current_row = 2
    data_start_row = current_row
    
    for dim_name, result in cognitive_results.items():
        worksheet.write(current_row, 0, dim_name)
        worksheet.write(current_row, 1, result['dominant_type'])
        worksheet.write(current_row, 2, result['percentage'])
        worksheet.write(current_row, 3, result['total_questions'])
        current_row += 1
    
    for dim_name, result in conative_results.items():
        worksheet.write(current_row, 0, dim_name)
        worksheet.write(current_row, 1, result['dominant_type'])
        worksheet.write(current_row, 2, result['percentage'])
        worksheet.write(current_row, 3, result['total_questions'])
        current_row += 1
    
    for dim_name, result in semantic_results.items():
        worksheet.write(current_row, 0, dim_name)
        worksheet.write(current_row, 1, result['dominant_type'])
        worksheet.write(current_row, 2, result['percentage'])
        worksheet.write(current_row, 3, result['total_questions'])
        current_row += 1
    
    for dim_name, result in emotional_results.items():
        worksheet.write(current_row, 0, dim_name)
        worksheet.write(current_row, 1, result['dominant_type'])
        worksheet.write(current_row, 2, result['percentage'])
        worksheet.write(current_row, 3, result['total_questions'])
        current_row += 1
    
    data_end_row = current_row - 1
    
    if data_end_row >= data_start_row:
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': 'Strength (%)',
            'categories': ['Summary', data_start_row, 0, data_end_row, 0],
            'values': ['Summary', data_start_row, 2, data_end_row, 2],
            'fill': {'color': '#4CAF50'},
            'data_labels': {'value': True}
        })
        chart.set_title({'name': 'All Dimensions - Complete Profile'})
        chart.set_x_axis({'name': 'Strength (%)', 'min': 0, 'max': 100})
        chart.set_y_axis({'name': 'Dimension'})
        
        num_dimensions = len(cognitive_results) + len(conative_results) + len(semantic_results) + len(emotional_results)
        chart_height = max(600, num_dimensions * 20)
        chart.set_size({'width': 720, 'height': chart_height})
        chart.set_legend({'position': 'none'})
        worksheet.insert_chart(1, 5, chart)
    
    # 2. ACCUMULATED CHARTS PAGES (right after Summary)
    if cognitive_results:
        _create_charts_page(workbook, cognitive_results, 'Cognitive Charts', header_format)
    if conative_results:
        _create_charts_page(workbook, conative_results, 'Conative Charts', header_format)
    if semantic_results:
        _create_charts_page(workbook, semantic_results, 'Semantic Charts', header_format)
    if emotional_results:
        _create_charts_page(workbook, emotional_results, 'Emotional Charts', header_format)
    
    # 3. PERSONAL PROFILE (after accumulated charts)
    profile_text = generate_personal_profile(cognitive_results, conative_results, 
                                            semantic_results, emotional_results)
    
    profile_ws = workbook.add_worksheet('Personal Profile')
    profile_ws.set_column('A:A', 120)
    
    # Title
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'bg_color': '#2196F3',
        'font_color': 'white',
        'align': 'center'
    })
    profile_ws.merge_range('A1:A2', f'PERSONAL PROFILE - {client_name}', title_format)
    
    # Profile content
    profile_ws.write(3, 0, profile_text, text_wrap_format)
    
    # 4. INDIVIDUAL DIMENSION SHEETS
    _process_dimension_sheets(workbook, cognitive_responses, cognitive_dimensions, cognitive_results, header_format)
    _process_dimension_sheets(workbook, conative_responses, conative_dimensions, conative_results, header_format)
    _process_dimension_sheets(workbook, semantic_responses, semantic_dimensions, semantic_results, header_format)
    _process_dimension_sheets(workbook, emotional_responses, emotional_dimensions, emotional_results, header_format)
    
    # 5. INFO SHEET
    info_ws = workbook.add_worksheet('Info')
    info_ws.set_column('A:A', 30)
    info_ws.set_column('B:B', 40)
    
    info_data = [
        ['Field', 'Value'],
        ['Client Name', client_name],
        ['Assessment Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
        ['Cognitive Dimensions', len(cognitive_results)],
        ['Cognitive Questions', sum(r['total_questions'] for r in cognitive_results.values()) if cognitive_results else 0],
        ['Conative Dimensions', len(conative_results)],
        ['Conative Questions', sum(r['total_questions'] for r in conative_results.values()) if conative_results else 0],
        ['Semantic Dimensions', len(semantic_results)],
        ['Semantic Questions', sum(r['total_questions'] for r in semantic_results.values()) if semantic_results else 0],
        ['Emotional Dimensions', len(emotional_results)],
        ['Emotional Questions', sum(r['total_questions'] for r in emotional_results.values()) if emotional_results else 0],
        ['Total Dimensions', len(cognitive_results) + len(conative_results) + len(semantic_results) + len(emotional_results)],
        ['Total Questions', sum([
            sum(r['total_questions'] for r in cognitive_results.values()) if cognitive_results else 0,
            sum(r['total_questions'] for r in conative_results.values()) if conative_results else 0,
            sum(r['total_questions'] for r in semantic_results.values()) if semantic_results else 0,
            sum(r['total_questions'] for r in emotional_results.values()) if emotional_results else 0
        ])]
    ]
    
    for row_idx, row_data in enumerate(info_data):
        for col_idx, value in enumerate(row_data):
            if row_idx == 0:
                info_ws.write(row_idx, col_idx, value, header_format)
            else:
                info_ws.write(row_idx, col_idx, value)
    
    workbook.close()
    buffer.seek(0)
    return buffer

def _create_charts_page(workbook, results, sheet_name, header_format):
    """Create accumulated charts page"""
    worksheet = workbook.add_worksheet(sheet_name)
    
    chart_num = 0
    for dim_name, result in results.items():
        clean_name = re.sub(r'^\d+\.', '', dim_name).strip()
        clean_name = clean_name.split(':')[0].strip()
        
        data_col = 26 + (chart_num * 3)
        worksheet.write(1, data_col, 'Type', header_format)
        worksheet.write(1, data_col + 1, 'Percentage', header_format)
        
        data_row = 2
        for type_name, count in result['all_scores'].items():
            percentage = (count / result['total_questions']) * 100
            worksheet.write(data_row, data_col, type_name)
            worksheet.write(data_row, data_col + 1, percentage)
            data_row += 1
        
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': 'Percentage',
            'categories': [sheet_name, 2, data_col, data_row - 1, data_col],
            'values': [sheet_name, 2, data_col + 1, data_row - 1, data_col + 1],
            'fill': {'color': '#2196F3'},
            'data_labels': {'value': True}
        })
        chart.set_title({'name': clean_name})
        chart.set_x_axis({'name': 'Percentage (%)', 'min': 0, 'max': 100})
        chart.set_y_axis({'name': 'Type'})
        chart.set_size({'width': 480, 'height': 300})
        chart.set_legend({'position': 'none'})
        
        chart_col = (chart_num % 2) * 8
        chart_row = (chart_num // 2) * 18 + 1
        worksheet.insert_chart(chart_row, chart_col, chart)
        
        chart_num += 1

def _process_dimension_sheets(workbook, responses, dimensions, results, header_format):
    """Create individual dimension sheets"""
    for dim_idx, dimension in enumerate(dimensions):
        if dim_idx not in responses:
            continue
        
        dim_name = dimension['name']
        clean_name = re.sub(r'^\d+\.', '', dim_name).strip()
        clean_name = clean_name.split(':')[0].strip()
        sheet_name = clean_name[:31]
        
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 60)
        worksheet.set_column('C:C', 60)
        
        headers = ['Question #', 'Question', 'Answer']
        for col, header in enumerate(headers):
            worksheet.write(1, col, header, header_format)
        
        row_idx = 2
        for q_idx, question in enumerate(dimension['questions'], 1):
            answer_key = responses[dim_idx].get(q_idx)
            answer_text = ''
            
            if not answer_key:
                answer_key = 'N/A'
            elif answer_key != 'N/A':
                try:
                    option_idx = ord(answer_key) - ord('a')
                    if 0 <= option_idx < len(question['options']):
                        answer_text = question['options'][option_idx]
                except (TypeError, ValueError):
                    answer_key = 'N/A'
                    answer_text = ''
            
            worksheet.write(row_idx, 0, q_idx)
            worksheet.write(row_idx, 1, question['text'])
            worksheet.write(row_idx, 2, f"{answer_key}) {answer_text}" if answer_text else answer_key)
            row_idx += 1
        
        if dim_name in results:
            result = results[dim_name]
            
            chart_start_row = row_idx + 3
            worksheet.write(chart_start_row, 0, 'Type', header_format)
            worksheet.write(chart_start_row, 1, 'Count', header_format)
            worksheet.write(chart_start_row, 2, 'Percentage', header_format)
            
            chart_data_row = chart_start_row + 1
            for type_name, count in result['all_scores'].items():
                percentage = (count / result['total_questions']) * 100
                worksheet.write(chart_data_row, 0, type_name)
                worksheet.write(chart_data_row, 1, count)
                worksheet.write(chart_data_row, 2, percentage)
                chart_data_row += 1
            
            chart = workbook.add_chart({'type': 'bar'})
            chart.add_series({
                'name': 'Percentage',
                'categories': [sheet_name, chart_start_row + 1, 0, chart_data_row - 1, 0],
                'values': [sheet_name, chart_start_row + 1, 2, chart_data_row - 1, 2],
                'fill': {'color': '#2196F3'},
                'data_labels': {'value': True}
            })
            chart.set_title({'name': f'{clean_name} - Distribution'})
            chart.set_x_axis({'name': 'Percentage (%)', 'min': 0, 'max': 100})
            chart.set_y_axis({'name': 'Type'})
            chart.set_size({'width': 480, 'height': 300})
            chart.set_legend({'position': 'none'})
            worksheet.insert_chart(chart_start_row, 4, chart)

def main():
    st.set_page_config(
        page_title="NLP Complete Assessment", 
        page_icon="🧠", 
        layout="wide"
    )
    
    st.markdown("""
        <style>
        .stButton button {
            width: 100%;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("🧠 NLP Complete Assessment - 4 Dimensions")
    st.markdown("---")
    
    if 'cognitive_dimensions' not in st.session_state:
        with open('Cognitive.txt', 'r', encoding='utf-8') as f:
            cognitive_content = f.read()
        st.session_state.cognitive_dimensions = parse_assessment_file(cognitive_content)
    
    if 'conative_dimensions' not in st.session_state:
        with open('Conative.txt', 'r', encoding='utf-8') as f:
            conative_content = f.read()
        st.session_state.conative_dimensions = parse_assessment_file(conative_content)
    
    if 'semantic_dimensions' not in st.session_state:
        with open('Semantic.txt', 'r', encoding='utf-8') as f:
            semantic_content = f.read()
        st.session_state.semantic_dimensions = parse_assessment_file(semantic_content)
    
    if 'emotional_dimensions' not in st.session_state:
        with open('Emotional.txt', 'r', encoding='utf-8') as f:
            emotional_content = f.read()
        st.session_state.emotional_dimensions = parse_assessment_file(emotional_content)
    
    if 'client_name' not in st.session_state:
        st.session_state.client_name = ""
    if 'current_section' not in st.session_state:
        st.session_state.current_section = 0
    if 'current_dimension' not in st.session_state:
        st.session_state.current_dimension = 0
    if 'cognitive_responses' not in st.session_state:
        st.session_state.cognitive_responses = {}
    if 'conative_responses' not in st.session_state:
        st.session_state.conative_responses = {}
    if 'semantic_responses' not in st.session_state:
        st.session_state.semantic_responses = {}
    if 'emotional_responses' not in st.session_state:
        st.session_state.emotional_responses = {}
    if 'assessment_complete' not in st.session_state:
        st.session_state.assessment_complete = False
    
    sections = [
        ('cognitive', st.session_state.cognitive_dimensions, st.session_state.cognitive_responses, "Cognitive (Thinking)", "Part 1 of 4"),
        ('conative', st.session_state.conative_dimensions, st.session_state.conative_responses, "Conative (Choosing)", "Part 2 of 4"),
        ('semantic', st.session_state.semantic_dimensions, st.session_state.semantic_responses, "Semantic (Meta)", "Part 3 of 4"),
        ('emotional', st.session_state.emotional_dimensions, st.session_state.emotional_responses, "Emotional (Feeling)", "Part 4 of 4")
    ]
    
    section_key, dimensions, responses_dict, section_name, section_progress = sections[st.session_state.current_section]
    
    if not st.session_state.client_name:
        st.markdown("### Welcome to the Complete NLP Assessment")
        st.write("This assessment covers 4 dimensions: Cognitive, Conative, Semantic, and Emotional")
        st.write("")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            name = st.text_input("Your Name:", key="name_input")
        
        st.write("")
        
        col1, col2, col3 = st.columns([1, 1, 2])
        with col2:
            if st.button("Start Assessment", type="primary", use_container_width=True):
                if name and name.strip():
                    st.session_state.client_name = name.strip()
                    st.rerun()
                else:
                    st.error("Please enter your name to continue.")
    
    elif not st.session_state.assessment_complete:
        dim_idx = st.session_state.current_dimension
        
        if dim_idx >= len(dimensions):
            if st.session_state.current_section < 3:
                st.session_state.current_section += 1
                st.session_state.current_dimension = 0
                st.rerun()
            else:
                st.session_state.assessment_complete = True
                st.rerun()
            return
        
        dimension = dimensions[dim_idx]
        
        st.markdown(f"### 📋 {section_progress}: {section_name}")
        progress = (dim_idx + 1) / len(dimensions)
        st.progress(progress)
        st.caption(f"Dimension {dim_idx + 1} of {len(dimensions)} in {section_name}")
        
        st.markdown("---")
        st.markdown(f"#### {dimension['name']}")
        st.write("")
        
        if dim_idx not in responses_dict:
            responses_dict[dim_idx] = {}
        
        for q_idx, question in enumerate(dimension['questions'], 1):
            st.markdown(f"**Question {q_idx}:** {question['text']}")
            
            previous_answer = responses_dict[dim_idx].get(q_idx)
            
            if previous_answer:
                default_index = ord(previous_answer) - ord('a')
            else:
                default_index = 0
            
            answer = st.radio(
                f"Select your answer for Question {q_idx}:",
                options=[chr(97 + i) for i in range(len(question['options']))],
                format_func=lambda x, opts=question['options']: f"{x}) {opts[ord(x) - ord('a')]}",
                key=f"q_{st.session_state.current_section}_{dim_idx}_{q_idx}",
                index=default_index,
                label_visibility="collapsed"
            )
            
            if answer is not None:
                responses_dict[dim_idx][q_idx] = answer
            
            st.write("")
        
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if dim_idx > 0:
                if st.button("⬅️ Previous Dimension", type="secondary", use_container_width=True):
                    st.session_state.current_dimension -= 1
                    st.rerun()
            elif st.session_state.current_section > 0:
                if st.button("⬅️ Previous Section", type="secondary", use_container_width=True):
                    st.session_state.current_section -= 1
                    prev_section = sections[st.session_state.current_section]
                    st.session_state.current_dimension = len(prev_section[1]) - 1
                    st.rerun()
        
        with col3:
            if dim_idx < len(dimensions) - 1:
                if st.button("Next Dimension ➡️", type="primary", use_container_width=True):
                    st.session_state.current_dimension += 1
                    st.rerun()
            else:
                if st.session_state.current_section < 3:
                    if st.button("Continue to Next Part ➡️", type="primary", use_container_width=True):
                        st.session_state.current_section += 1
                        st.session_state.current_dimension = 0
                        st.rerun()
                else:
                    if st.button("Finish Assessment ✅", type="primary", use_container_width=True):
                        st.session_state.assessment_complete = True
                        st.rerun()
    
    else:
        st.success("✅ Assessment Complete!")
        st.markdown("---")
        
        cognitive_results = calculate_results(st.session_state.cognitive_responses, st.session_state.cognitive_dimensions)
        conative_results = calculate_results(st.session_state.conative_responses, st.session_state.conative_dimensions)
        semantic_results = calculate_results(st.session_state.semantic_responses, st.session_state.semantic_dimensions)
        emotional_results = calculate_results(st.session_state.emotional_responses, st.session_state.emotional_dimensions)
        
        st.markdown("### Your Complete Profile Summary")
        st.write("")
        
        st.markdown("## 🧠 Cognitive (Thinking) - 18 Dimensions")
        for dim_name, result in cognitive_results.items():
            with st.expander(f"📊 {dim_name}", expanded=False):
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown(f"**Dominant Type:** {result['dominant_type']}")
                    st.markdown(f"**Strength:** {result['percentage']:.1f}%")
                    st.markdown(f"**Questions Answered:** {result['total_questions']}")
                with col2:
                    st.markdown("**Score Breakdown:**")
                    for type_name, count in result['all_scores'].items():
                        percentage = (count / result['total_questions']) * 100
                        st.markdown(f"• {type_name}: {count} ({percentage:.0f}%)")
        
        st.markdown("---")
        
        st.markdown("## 🎯 Conative (Choosing) - 14 Dimensions")
        for dim_name, result in conative_results.items():
            with st.expander(f"📊 {dim_name}", expanded=False):
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown(f"**Dominant Type:** {result['dominant_type']}")
                    st.markdown(f"**Strength:** {result['percentage']:.1f}%")
                    st.markdown(f"**Questions Answered:** {result['total_questions']}")
                with col2:
                    st.markdown("**Score Breakdown:**")
                    for type_name, count in result['all_scores'].items():
                        percentage = (count / result['total_questions']) * 100
                        st.markdown(f"• {type_name}: {count} ({percentage:.0f}%)")
        
        st.markdown("---")
        
        st.markdown("## 🔮 Semantic (Meta) - 13 Dimensions")
        for dim_name, result in semantic_results.items():
            with st.expander(f"📊 {dim_name}", expanded=False):
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown(f"**Dominant Type:** {result['dominant_type']}")
                    st.markdown(f"**Strength:** {result['percentage']:.1f}%")
                    st.markdown(f"**Questions Answered:** {result['total_questions']}")
                with col2:
                    st.markdown("**Score Breakdown:**")
                    for type_name, count in result['all_scores'].items():
                        percentage = (count / result['total_questions']) * 100
                        st.markdown(f"• {type_name}: {count} ({percentage:.0f}%)")
        
        st.markdown("---")
        
        st.markdown("## ❤️ Emotional (Feeling) - 15 Dimensions")
        for dim_name, result in emotional_results.items():
            with st.expander(f"📊 {dim_name}", expanded=False):
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown(f"**Dominant Type:** {result['dominant_type']}")
                    st.markdown(f"**Strength:** {result['percentage']:.1f}%")
                    st.markdown(f"**Questions Answered:** {result['total_questions']}")
                with col2:
                    st.markdown("**Score Breakdown:**")
                    for type_name, count in result['all_scores'].items():
                        percentage = (count / result['total_questions']) * 100
                        st.markdown(f"• {type_name}: {count} ({percentage:.0f}%)")
        
        st.markdown("---")
        st.markdown("### Download Your Complete Results")
        
        excel_buffer = create_excel_report(
            st.session_state.client_name,
            st.session_state.cognitive_responses,
            st.session_state.cognitive_dimensions,
            st.session_state.conative_responses,
            st.session_state.conative_dimensions,
            st.session_state.semantic_responses,
            st.session_state.semantic_dimensions,
            st.session_state.emotional_responses,
            st.session_state.emotional_dimensions
        )
        
        filename = f"NLP_Complete_Assessment_{st.session_state.client_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        st.download_button(
            label="📥 Download Complete Excel Report (60 Dimensions + Profile)",
            data=excel_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        st.write("")
        
        if st.button("Start New Assessment", type="secondary"):
            st.session_state.client_name = ""
            st.session_state.current_section = 0
            st.session_state.current_dimension = 0
            st.session_state.cognitive_responses = {}
            st.session_state.conative_responses = {}
            st.session_state.semantic_responses = {}
            st.session_state.emotional_responses = {}
            st.session_state.assessment_complete = False
            st.rerun()

if __name__ == "__main__":
    main()
