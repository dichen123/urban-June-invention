import pandas as pd
import re

def convert_final_questions(input_file, output_file):
    df = pd.read_excel(input_file, engine='openpyxl', sheet_name=0, header=0)
    error_log = []
    
    # 列名标准化
    df = df.rename(columns={
        df.columns[0]: '序号',
        df.columns[1]: '题目和选项',
        df.columns[2]: '答案'
    })
    
    df['填空题'] = ''
    df['答案文本'] = ''

    # 超强选项解析正则
    option_pattern = re.compile(r'''
        ([A-D])                  # 选项字母
        [、;；\.]?               # 分隔符（允许缺失）
        \s*                      
        (                       
            (?!\s*[A-D][、;；.]) # 排除下一个选项
            [^\sA-D]+           # 内容
        )
    ''', re.VERBOSE | re.UNICODE)

    for index, row in df.iterrows():
        full_text = str(row['题目和选项'])
        raw_answer = re.sub(r'[^A-D]', '', str(row['答案']).upper())
        question_part = ""
        
        try:
            # 切割题干和选项
            split_result = re.split(r'(?=([A-D])[、;；.\s])', full_text, maxsplit=1)
            question_part = split_result[0].replace("（）", "____").strip()
            options_text = ''.join(split_result[1:]) if len(split_result)>1 else ""

            # 解析选项
            options = {}
            matches = option_pattern.finditer(options_text)
            for match in matches:
                key = match.group(1)
                value = match.group(2).strip()
                options[key] = value

            # 答案验证
            if not options:
                raise ValueError("无有效选项")
                
            answer_letters = list(raw_answer)
            valid_answers = []
            for letter in answer_letters:
                if letter in options:
                    valid_answers.append(options[letter])
                else:
                    raise ValueError(f"选项{letter}不存在 | 实际选项: {list(options.keys())}")

            # 生成填空题
            blank = "____" if len(valid_answers)==1 else f"____({len(valid_answers)}项)"
            df.at[index, '填空题'] = question_part + blank
            df.at[index, '答案文本'] = "；".join(valid_answers)

        except Exception as e:
            error_log.append(f"第{index+1}行错误：{str(e)}")
            df.at[index, '填空题'] = "【错误】" + str(full_text)[:50]
            df.at[index, '答案文本'] = "【解析失败】"

    # 输出结果
    df = df[['序号', '填空题', '答案文本']].rename(columns={'答案文本': '答案'})
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    if error_log:
        print("关键错误需人工核对：\n" + "\n".join(error_log[:10]))  # 仅显示前10个错误
    else:
        print("转换成功！")

convert_final_questions('input.xlsx', 'output.xlsx')