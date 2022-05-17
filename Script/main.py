from datetime import datetime
import os
import pandas as pd
import set_env
import generate_extract
import vbaformatter
import SC_Zip
import SC_Email

FMT = '%m/%d/%Y_%H:%M:%S'
now = datetime.now()
print("Start Time")
print(now.strftime(FMT))

child_dir = now.strftime("%m%d%Y_%H%M%S")
os.mkdir('C:/BAN_Extraction_Tool_V2.0/Output/%s' % child_dir)
ConfigFile = "Configuration.xlsm"
configuration_file = r'C:/BAN_Extraction_Tool_V2.0/%s' % ConfigFile
absolute_config_path = os.path.abspath(r'%s' % configuration_file)
ensemble_file_name = "Converted_Ensemble_Extract_Money_Map_%s.xlsx" % now.strftime("%m%d%Y_%H%M")
ensemble_file_path = r'C:/BAN_Extraction_Tool_V2.0/Output/%s/%s' % (child_dir, ensemble_file_name)
absolute_ensemble_file_path = os.path.abspath(r'%s' % ensemble_file_path)
metro_file_name = "Converted_Metro_Extract_Money_Map_%s.xlsx" % now.strftime("%m%d%Y_%H%M")
metro_file_path = r'C:/BAN_Extraction_Tool_V2.0/Output/%s/%s' % (child_dir, metro_file_name)
absolute_metro_file_path = os.path.abspath(r'%s' % metro_file_path)
df = pd.read_excel(configuration_file, sheet_name="Input")
en_BAN = []
sa_BAN = []
ens_env = df.iloc[0, 5]
mag_env = df.iloc[1, 5]
to_email_id = df.fillna('').iloc[0, 7]
print(to_email_id)
for index in range(len(df)):
    i = index + 2
    s_ban = df.iloc[index, 0]
    m_ban = df.iloc[index, 1]
    en_BAN.append(s_ban)
    sa_BAN.append(m_ban)

ban_df = pd.DataFrame(list(zip(en_BAN, sa_BAN)), columns=['LY BAN', 'LM BAN'])
print(ban_df)
converted_e_ban = [str(element) for element in en_BAN]
converted_m_ban = [str(element) for element in sa_BAN]
e_BAN = ", ".join(converted_e_ban).strip()
m_BAN = ", ".join(converted_m_ban).strip()
if len(e_BAN) > 0 and len(m_BAN) > 0:
    en_conn = set_env.ensemble_env(file=configuration_file, env=ens_env)
    sa_conn = set_env.samson_env(file=configuration_file, env=mag_env)
    print("Please wait Extracts is getting generated................")
    generate_extract.data_extract(file=configuration_file, ban1=e_BAN, ban2=m_BAN, path1=ensemble_file_path,
                                  path2=metro_file_path, env1=en_conn, env2=sa_conn)
    work_book = [ensemble_file_name, metro_file_name]
    absolute_work_book = [absolute_ensemble_file_path, absolute_metro_file_path]
    vbaformatter.formatting_file(ab_path=absolute_config_path, wb_name=work_book, wb_name_ab_path=absolute_work_book,
                                 file1=ConfigFile, file2=ensemble_file_name, file3=metro_file_name)
else:
    print("Please provide BAN to be extracted in config file")

zip_file_path = "C:/BAN_Extraction_Tool_V2.0/Output/%s" % child_dir
SC_Zip.sc_zip(zip_file_path)
email_file = r'%s.zip' % zip_file_path
absolutezip_file_namepath = os.path.abspath(r'%s' % email_file)
SC_Email.email_file(email=to_email_id, sub=child_dir, abzipfilepath=absolutezip_file_namepath)

now = datetime.now()
print("End Time")
print(now.strftime(FMT))
