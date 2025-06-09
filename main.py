import xlwings as xw
from xlwings.constants import LineStyle
import os

FILE = "score.xlsx"
SHEET = "Sheet1"
FR_TO = "A4:N668"

RESULT = "result.xlsx"

txt_tin_ = "Tin Học (Chuyên)"
txt_toan_ = "Toán (Chuyên)"
txt_anh_ = "Tiếng Anh (Chuyên)"
txt_van_ = "Ngữ Văn (Chuyên)"
txt_su_ = "Lịch Sử (Chuyên)"
txt_dia_ = "Địa Lý (Chuyên)"
txt_sinh_ = "Sinh Học (Chuyên)"
txt_phap_ = "Tiếng Pháp (Chuyên)"
txt_li_ = "Vật Lí (Chuyên)"
txt_hoa_ = "Hóa Học (Chuyên)"

hs_ = 35

toan_ = 2
anh_ = 2
van_ = 2
chuyen_ = 4

sum_p = 34
toan_p = 4
anh_p = 8
van_p = 4
anh_cp = 10

toan_s = 5
van_s = 6
anh_s = 7
chuyen_s1 = 9
chuyen_s2 = 11
mon_s1 = 8
mon_s2 = 10

sum1_s = 12
sum2_s = 13

headers = ["SBD", "Họ và tên", "Giới tính", "Học sinh trường", "Toán", "Văn", "Anh", "Nguyện vọng 1", "Điểm NV1", "Nguyện vọng 2", "Điểm NV2", "Tổng NV1", "Tổng NV2", "Tổng điểm chuyên"]

def num(var):
    return isinstance(var, (int, float))

def safe_score(score):
    if score == "VT" or score is None:
        return 0
    try:
        return float(score)
    except (ValueError, TypeError):
        return 0

def toi_thieu_p(toan, anh, van, chuyen, sum):
    toan = safe_score(toan)
    anh = safe_score(anh)
    van = safe_score(van)
    chuyen = safe_score(chuyen)
    sum = safe_score(sum)
    
    if(toan >= toan_p and anh >= anh_p and van >= van_p and chuyen >= chuyen_ and sum >= sum_p):
        return True
    return False

def toi_thieu(toan, anh, van, chuyen):
    toan = safe_score(toan)
    anh = safe_score(anh)
    van = safe_score(van)
    chuyen = safe_score(chuyen)
    
    if(toan >= toan_ and anh >= anh_ and van >= van_ and chuyen >= chuyen_):
        return True
    return False

def can_admit(student, subject, preference_num):
    if preference_num == 1:
        chuyen_score = safe_score(student[chuyen_s1])
        total_score = safe_score(student[sum1_s])
        subject_col = mon_s1
    else:
        chuyen_score = safe_score(student[chuyen_s2])
        total_score = safe_score(student[sum2_s])
        subject_col = mon_s2
    
    if student[subject_col] != subject:
        return False, 0, 0
    
    if not toi_thieu(student[toan_s], student[anh_s], student[van_s], chuyen_score):
        return False, 0, 0
    
    return True, chuyen_score, total_score

def solve_competitive_admission(data):
    subjects = [txt_tin_, txt_toan_, txt_anh_, txt_van_, txt_su_, 
               txt_dia_, txt_sinh_, txt_phap_, txt_li_, txt_hoa_]
    
    admitted = {subject: [] for subject in subjects}
    
    admitted_students = set()
    
    first_pref_candidates = []
    for student in data:
        if not student or len(student) < max(sum1_s, sum2_s) + 1:
            continue
            
        student_id = student[0]
        if student_id is None:
            continue
            
        first_subject = student[mon_s1] if len(student) > mon_s1 and student[mon_s1] in subjects else None
        
        if first_subject:
            can_admit_first, chuyen1, total1 = can_admit(student, first_subject, 1)
            
            if can_admit_first:
                student_record = student[:] + [first_subject, chuyen1, total1, 1]
                first_pref_candidates.append((first_subject, student_record, total1))
    
    subject_first_candidates = {subject: [] for subject in subjects}
    for subject, student_record, total_score in first_pref_candidates:
        subject_first_candidates[subject].append((student_record, total_score))
    
    for subject in subjects:
        candidates = sorted(subject_first_candidates[subject], key=lambda x: x[1], reverse=True)
        for student_record, _ in candidates[:hs_]:
            admitted[subject].append(student_record)
            admitted_students.add(student_record[0])
    
    second_pref_candidates = []
    for student in data:
        if not student or len(student) < max(sum1_s, sum2_s) + 1:
            continue
            
        student_id = student[0]
        if student_id is None or student_id in admitted_students:
            continue
            
        second_subject = student[mon_s2] if len(student) > mon_s2 and student[mon_s2] in subjects else None
        
        if second_subject:
            can_admit_second, chuyen2, total2 = can_admit(student, second_subject, 2)
            
            if can_admit_second:
                student_record = student[:] + [second_subject, chuyen2, total2, 2]
                second_pref_candidates.append((second_subject, student_record, total2))
    
    second_pref_candidates.sort(key=lambda x: x[2], reverse=True)
    
    for subject, student_record, total_score in second_pref_candidates:
        student_id = student_record[0]
        
        if student_id in admitted_students:
            continue
        
        current_admitted = admitted[subject]
        
        if len(current_admitted) < hs_:
            admitted[subject].append(student_record)
            admitted_students.add(student_id)
        else:
            min_score = min(safe_score(s[-2]) for s in current_admitted)
            
            if total_score > min_score:
                for i, admitted_student in enumerate(current_admitted):
                    if safe_score(admitted_student[-2]) == min_score:
                        removed_student = current_admitted.pop(i)
                        admitted_students.remove(removed_student[0])
                        break
                
                admitted[subject].append(student_record)
                admitted_students.add(student_id)
    
    for subject in subjects:
        admitted[subject].sort(key=lambda x: safe_score(x[-2]), reverse=True)
        admitted[subject] = admitted[subject][:hs_]
    
    return (admitted[txt_tin_], admitted[txt_toan_], admitted[txt_anh_], 
            admitted[txt_van_], admitted[txt_su_], admitted[txt_dia_], 
            admitted[txt_sinh_], admitted[txt_phap_], admitted[txt_li_], 
            admitted[txt_hoa_])

def safe_sort_key(student, score_index):
    try:
        score = student[score_index]
        return safe_score(score)
    except (ValueError, TypeError, IndexError):
        return 0.0

def solve_p(phap, admitted_to_english, data):
    english_students = [s for s in data if s[mon_s1] == txt_anh_]
    english_students = sorted(english_students, 
                             key=lambda x: safe_score(x[sum1_s]), 
                             reverse=True)
    
    predict = []
    admitted_english_ids = {s[0] for s in admitted_to_english}
    admitted_phap_ids = {s[0] for s in phap}
    
    for student in english_students:
        student_id = student[0]
        if (student_id not in admitted_english_ids and 
            student_id not in admitted_phap_ids and
            toi_thieu_p(student[toan_s], student[anh_s], student[van_s], 
                        student[chuyen_s1], student[sum1_s])):
            
            student_record = student[:] + [txt_phap_, safe_score(student[chuyen_s1]), safe_score(student[sum1_s]), "Chuyển từ Anh"]
            predict.append(student_record)
    
    return sorted(predict, key=lambda x: safe_score(x[-2]), reverse=True)

def write(file, name_s, data):
    if os.path.exists(file) == False:
      wb = xw.Book()
      wb.save(file)
      wb.close()
      
    wb = xw.Book(file)
    if name_s not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(name=name_s)
        
    sheet = wb.sheets[name_s]
    for row_index, row in enumerate(data):
      for col_index, value in enumerate(row):
        cell = sheet.cells(row_index + 1, col_index + 1)
        cell.value = value
    wb.save()

def head(file, headers):
    wb = xw.Book(file)
    for sheet in wb.sheets:
        sheet.api.Rows("1:1").Insert()
        for i, header in enumerate(headers):
            sheet.cells(1, i + 1).value = header
    wb.save()
    
def fit(file):
    wb = xw.Book(file)
    for sheet in wb.sheets:
        sheet.autofit('c')
        sheet.autofit('r')
        used_range = sheet.used_range
        used_range.api.Borders.LineStyle = LineStyle.xlContinuous
    wb.save()

def rm_cols(row):
    n = len(row)
    drop = {0, n-1, n-3, n-4}
    return [v for i, v in enumerate(row) if i not in drop]
    
def rm_first_col(arr):
    return [rm_cols(r) for r in arr]

wb = xw.Book(FILE)
sheet = wb.sheets[SHEET]
data = sheet.range(FR_TO).value
wb.close()

tin, toan, anh, van, su, dia, sinh, phap, li, hoa = solve_competitive_admission(data)

phap_predict = solve_p(phap, anh, data)

tin_output = rm_first_col(tin)
toan_output = rm_first_col(toan)
anh_output = rm_first_col(anh)
van_output = rm_first_col(van)
su_output = rm_first_col(su)
dia_output = rm_first_col(dia)
sinh_output = rm_first_col(sinh)
phap_output = rm_first_col(phap)
li_output = rm_first_col(li)
hoa_output = rm_first_col(hoa)
phap_predict_output = rm_first_col(phap_predict)

write(RESULT, 'Toan', toan_output)
write(RESULT, 'Tin', tin_output)
write(RESULT, 'Van', van_output)
write(RESULT, 'Anh', anh_output)
write(RESULT, 'Su', su_output)
write(RESULT, 'Dia', dia_output)
write(RESULT, 'sinh', sinh_output)
write(RESULT, 'phap', phap_output)
write(RESULT, 'phap_anh', phap_predict_output)
write(RESULT, 'hoa', hoa_output)
write(RESULT, 'ly', li_output)

head(RESULT, headers)
fit(RESULT)
