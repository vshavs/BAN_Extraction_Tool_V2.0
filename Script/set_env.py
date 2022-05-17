import cx_Oracle
import pandas as pd
from sqlalchemy import create_engine

cx_Oracle.init_oracle_client(lib_dir=r"instantclient_19_10")


def ensemble_env(file, env):
    ensemble_env_df = pd.read_excel(file, sheet_name='Ensemble Env', index_col=None)
    for i, row in ensemble_env_df.iterrows():
        var_sl_no = (row['Env'])
        if var_sl_no == env:
            var_ensemble_hostname = (row['Ensemble_Hostname'])
            var_ensemble_port = (row['Ensemble_Port'])
            var_ensemble_sid = (row['Ensemble_SID'])
            var_ensemble_username = (row['Ensemble_Username'])
            var_ensemble_password = (row['Ensemble_Password'])
            ensemble_dsn_tns = cx_Oracle.makedsn(var_ensemble_hostname, var_ensemble_port, sid=var_ensemble_sid)
            en_conn = cx_Oracle.connect(user=var_ensemble_username, password=var_ensemble_password,
                                        dsn=ensemble_dsn_tns)
            return en_conn


def samson_env(file, env):
    samson_env_df = pd.read_excel(file, sheet_name='Magenta Env', index_col=None)
    for i, row in samson_env_df.iterrows():
        var_sl_no = (row['Env'])
        if var_sl_no == env:
            var_samson_hostname = (row['Samson_Hostname'])
            var_samson_port = (row['Samson_Port'])
            var_samson_sid = (row['Samson_SID'])
            var_samson_username = (row['Samson_Username'])
            var_samson_password = (row['Samson_Password'])
            samson_dsn_tns = cx_Oracle.makedsn(var_samson_hostname, var_samson_port, sid=var_samson_sid)
            sa_conn = cx_Oracle.connect(user=var_samson_username, password=var_samson_password,
                                        dsn=samson_dsn_tns)
            return sa_conn


