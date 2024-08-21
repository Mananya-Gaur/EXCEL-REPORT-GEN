# from flask import Flask, render_template, request, send_file
# import pandas as pd
# import os
# from io import BytesIO
# import tempfile
#
# app = Flask(__name__)
# app.secret_key = 'your_secret_key'  # Needed for session handling
#
#
# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'POST':
#         file1 = request.files['file1']
#         selected_date = pd.to_datetime(request.form['date'])  # Match the HTML form field name
#
#         # Process the uploaded files
#         df1 = pd.read_excel(file1)
#
#         # Combine and process the dataframes (example logic)
#         df_combined = df1
#         df_combined['Login Date'] = pd.to_datetime(df_combined['Login Date'], format='%d-%m-%Y')
#         filtered_df = df_combined[df_combined['Login Date'] == selected_date]
#
#         # Further processing and generating the output dataframe
#         output_df = pd.DataFrame(columns=[
#             'CCM', 'BM Name', 'Branches', 'No (Login)', 'Val (In Lacs) (Login)',
#             'No (Sanction)', 'Val (In Lacs) (Sanction)', 'No (Reject/Withdraw)',
#             'Val (In Lacs) (Reject/Withdraw)', 'No (Decision)', 'Val (In Lacs) (Decision)',
#             'No (Disbursed)', 'Val (In Lacs) (Disbursed)'
#         ])
#
#         grouped = filtered_df.groupby(['CCM', 'BM Name', 'Branches'])
#
#         for (ccm, bm_name, branches), group in grouped:
#             login_count = group['Lead ID (Synofin)'].nunique()
#             login_val = group['Request Amount'].sum() / 100000
#             sanction_count = group[group['Initial File Status (Credit)'] == 'Sanction'].shape[0]
#             sanction_val = group[group['Initial File Status (Credit)'] == 'Sanction']['Sanction Amount'].sum() / 100000
#             reject_count = group[group['Initial File Status (Credit)'] == 'Reject'].shape[0]
#             reject_val = group[group['Initial File Status (Credit)'] == 'Reject']['Request Amount'].sum() / 100000
#             decision_count = sanction_count + reject_count
#             decision_val = sanction_val + reject_val
#             disbursed_count = group[group['Disb. Date'].notnull()].shape[0]
#             disbursed_val = group[group['Disb. Date'].notnull()]['Sanction Amount'].sum() / 100000
#
#             output_df = output_df.append({
#                 'CCM': ccm,
#                 'BM Name': bm_name,
#                 'Branches': branches,
#                 'No (Login)': login_count,
#                 'Val (In Lacs) (Login)': login_val,
#                 'No (Sanction)': sanction_count,
#                 'Val (In Lacs) (Sanction)': sanction_val,
#                 'No (Reject/Withdraw)': reject_count,
#                 'Val (In Lacs) (Reject/Withdraw)': reject_val,
#                 'No (Decision)': decision_count,
#                 'Val (In Lacs) (Decision)': decision_val,
#                 'No (Disbursed)': disbursed_count,
#                 'Val (In Lacs) (Disbursed)': disbursed_val
#             }, ignore_index=True)
#
#         # Convert the dataframe to an HTML table for preview
#         output_html = output_df.to_html(classes='table table-striped', index=False)
#
#         # Save the output_df to a temporary file
#         temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
#         with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
#             output_df.to_excel(writer, index=False, sheet_name='Report')
#
#         # Store the path to pass it for downloading
#         temp_file_path = temp_file.name
#
#         return render_template('preview.html', table=output_html, temp_file_path=temp_file_path)
#
#     return render_template('index.html')
#
#
# @app.route('/download', methods=['POST'])
# def download():
#     temp_path = request.form['temp_file_path']
#
#     if not temp_path or not os.path.exists(temp_path):
#         return "File not found or the path is invalid", 400
#
#     # Send the file as an attachment
#     return send_file(temp_path, as_attachment=True, download_name='report.xlsx',
#                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#
#
# if __name__ == '__main__':
#     app.run(debug=True)

# from flask import Flask, render_template, request, send_file
# import pandas as pd
# import os
# from io import BytesIO
# import tempfile
#
# app = Flask(__name__)
# app.secret_key = 'your_secret_key'  # Needed for session handling
#
#
# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'POST':
#         file1 = request.files['file1']
#         selected_date = pd.to_datetime(request.form['date'])  # Match the HTML form field name
#
#         # Process the uploaded files
#         df1 = pd.read_excel(file1)
#
#         # Combine and process the dataframes (example logic)
#         df_combined = df1
#         df_combined['Login Date'] = pd.to_datetime(df_combined['Login Date'], format='%d-%m-%Y')
#         filtered_df = df_combined[df_combined['Login Date'] == selected_date]
#
#         # Further processing and generating the output dataframe
#         output_df = pd.DataFrame(columns=[
#             'CCM', 'No (Login)', 'Val (In Lacs) (Login)',
#             'No (Sanction)', 'Val (In Lacs) (Sanction)', 'No (Reject/Withdraw)',
#             'Val (In Lacs) (Reject/Withdraw)', 'No (Decision)', 'Val (In Lacs) (Decision)',
#             'No (Disbursed)', 'Val (In Lacs) (Disbursed)'
#         ])
#
#         grouped = filtered_df.groupby(['CCM'])
#
#         for (ccm), group in grouped:
#             login_count = group['Lead ID (Synofin)'].nunique()
#             login_val = group['Request Amount'].sum() / 100000
#             sanction_count = group[group['Initial File Status (Credit)'] == 'Sanction'].shape[0]
#             sanction_val = group[group['Initial File Status (Credit)'] == 'Sanction']['Sanction Amount'].sum() / 100000
#             reject_count = group[group['Initial File Status (Credit)'] == 'Reject'].shape[0]
#             reject_val = group[group['Initial File Status (Credit)'] == 'Reject']['Request Amount'].sum() / 100000
#             decision_count = sanction_count + reject_count
#             decision_val = sanction_val + reject_val
#             disbursed_count = group[group['Disb. Date'].notnull()].shape[0]
#             disbursed_val = group[group['Disb. Date'].notnull()]['Sanction Amount'].sum() / 100000
#
#             output_df = output_df.append({
#                 'CCM': ccm,
#                 'No (Login)': login_count,
#                 'Val (In Lacs) (Login)': login_val,
#                 'No (Sanction)': sanction_count,
#                 'Val (In Lacs) (Sanction)': sanction_val,
#                 'No (Reject/Withdraw)': reject_count,
#                 'Val (In Lacs) (Reject/Withdraw)': reject_val,
#                 'No (Decision)': decision_count,
#                 'Val (In Lacs) (Decision)': decision_val,
#                 'No (Disbursed)': disbursed_count,
#                 'Val (In Lacs) (Disbursed)': disbursed_val
#             }, ignore_index=True)
#
#         # Convert the dataframe to an HTML table for preview
#         output_html = output_df.to_html(classes='table table-striped', index=False)
#
#         # Save the output_df to a temporary file
#         temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
#         with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
#             output_df.to_excel(writer, index=False, sheet_name='Report')
#
#         # Store the path to pass it for downloading
#         temp_file_path = temp_file.name
#
#         return render_template('preview.html', table=output_html, temp_file_path=temp_file_path)
#
#     return render_template('index.html')
#
#
# @app.route('/download', methods=['POST'])
# def download():
#     temp_path = request.form['temp_file_path']
#
#     if not temp_path or not os.path.exists(temp_path):
#         return "File not found or the path is invalid", 400
#
#     # Send the file as an attachment
#     return send_file(temp_path, as_attachment=True, download_name='report.xlsx',
#                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#
#
# if __name__ == '__main__':
#     app.run(debug=True)

# from flask import Flask, render_template, request, send_file
# import pandas as pd
# import os
# from io import BytesIO
# import tempfile
#
# app = Flask(__name__)
# app.secret_key = 'your_secret_key'  # Needed for session handling
#
# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'POST':
#         file1 = request.files['file1']
#         selected_date = pd.to_datetime(request.form['date'])  # Match the HTML form field name
#
#         # Process the uploaded files
#         df1 = pd.read_excel(file1)
#
#         # Combine and process the dataframes (example logic)
#         df_combined = df1
#         df_combined['Login Date'] = pd.to_datetime(df_combined['Login Date'], format='%d-%m-%Y')
#         filtered_df = df_combined[df_combined['Login Date'] == selected_date]
#
#         # Further processing and generating the output dataframe
#         output_df = pd.DataFrame(columns=[
#             'CCM', 'No (Login)', 'Val (In Lacs) (Login)',
#             'No (Sanction)', 'Val (In Lacs) (Sanction)', 'No (Reject/Withdraw)',
#             'Val (In Lacs) (Reject/Withdraw)', 'No (Decision)', 'Val (In Lacs) (Decision)',
#             'No (Disbursed)', 'Val (In Lacs) (Disbursed)'
#         ])
#
#         grouped = filtered_df.groupby(['CCM'])
#
#         for (ccm), group in grouped:
#             login_count = group['Lead ID (Synofin)'].nunique()
#             login_val = group['Request Amount'].sum() / 100000
#             sanction_count = group[group['Initial File Status (Credit)'] == 'Sanction'].shape[0]
#             sanction_val = group[group['Initial File Status (Credit)'] == 'Sanction']['Sanction Amount'].sum() / 100000
#             reject_count = group[group['Initial File Status (Credit)'] == 'Reject'].shape[0]
#             reject_val = group[group['Initial File Status (Credit)'] == 'Reject']['Request Amount'].sum() / 100000
#             decision_count = sanction_count + reject_count
#             decision_val = sanction_val + reject_val
#             disbursed_count = group[group['Disb. Date'].notnull()].shape[0]
#             disbursed_val = group[group['Disb. Date'].notnull()]['Sanction Amount'].sum() / 100000
#
#             output_df = output_df.append({
#                 'CCM': ccm,
#                 'No (Login)': login_count,
#                 'Val (In Lacs) (Login)': login_val,
#                 'No (Sanction)': sanction_count,
#                 'Val (In Lacs) (Sanction)': sanction_val,
#                 'No (Reject/Withdraw)': reject_count,
#                 'Val (In Lacs) (Reject/Withdraw)': reject_val,
#                 'No (Decision)': decision_count,
#                 'Val (In Lacs) (Decision)': decision_val,
#                 'No (Disbursed)': disbursed_count,
#                 'Val (In Lacs) (Disbursed)': disbursed_val
#             }, ignore_index=True)
#
#         # Calculate totals
#         totals = {
#             'CCM': 'Total',
#             'No (Login)': output_df['No (Login)'].sum(),
#             'Val (In Lacs) (Login)': output_df['Val (In Lacs) (Login)'].sum(),
#             'No (Sanction)': output_df['No (Sanction)'].sum(),
#             'Val (In Lacs) (Sanction)': output_df['Val (In Lacs) (Sanction)'].sum(),
#             'No (Reject/Withdraw)': output_df['No (Reject/Withdraw)'].sum(),
#             'Val (In Lacs) (Reject/Withdraw)': output_df['Val (In Lacs) (Reject/Withdraw)'].sum(),
#             'No (Decision)': output_df['No (Decision)'].sum(),
#             'Val (In Lacs) (Decision)': output_df['Val (In Lacs) (Decision)'].sum(),
#             'No (Disbursed)': output_df['No (Disbursed)'].sum(),
#             'Val (In Lacs) (Disbursed)': output_df['Val (In Lacs) (Disbursed)'].sum()
#
#
#         }
#
#         # Append the totals row
#         output_df = output_df.append(totals, ignore_index=True)
#
#         # Convert the dataframe to an HTML table for preview
#         output_html = output_df.to_html(classes='table table-striped', index=False)
#
#         # Save the output_df to a temporary file
#         temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
#         with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
#             output_df.to_excel(writer, index=False, sheet_name='Report')
#
#         # Store the path to pass it for downloading
#         temp_file_path = temp_file.name
#
#         return render_template('preview.html', table=output_html, temp_file_path=temp_file_path)
#
#     return render_template('index.html')
#
#
# @app.route('/download', methods=['POST'])
# def download():
#     temp_path = request.form['temp_file_path']
#
#     if not temp_path or not os.path.exists(temp_path):
#         return "File not found or the path is invalid", 400
#
#     # Send the file as an attachment
#     return send_file(temp_path, as_attachment=True, download_name='report.xlsx',
#                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#
#
# if __name__ == '__main__':
#     app.run(debug=True)

from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from io import BytesIO
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Needed for session handling

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file1 = request.files['file1']
        selected_date = pd.to_datetime(request.form['date'])  # Get date input from the user

        # Process the uploaded file
        df = pd.read_excel(file1)

        # First DataFrame processing (from code1)
        df1 = df.copy()
        df1['Login Date'] = pd.to_datetime(df1['Login Date'], format='%d-%m-%Y')
        filtered_df = df1[df1['Login Date'] == selected_date]

        output_df1 = pd.DataFrame(columns=[
            'CCM', 'No (Login)', 'Val (In Lacs) (Login)',
            'No (Sanction)', 'Val (In Lacs) (Sanction)', 'No (Reject/Withdraw)',
            'Val (In Lacs) (Reject/Withdraw)', 'No (Decision)', 'Val (In Lacs) (Decision)',
            'No (Disbursed)', 'Val (In Lacs) (Disbursed)'
        ])

        grouped = filtered_df.groupby(['CCM'])

        for (ccm), group in grouped:
            login_count = group['Lead ID (Synofin)'].nunique()
            login_val = group['Request Amount'].sum() / 100000
            sanction_count = group[group['Initial File Status (Credit)'] == 'Sanction'].shape[0]
            sanction_val = group[group['Initial File Status (Credit)'] == 'Sanction']['Sanction Amount'].sum() / 100000
            reject_count = group[group['Initial File Status (Credit)'] == 'Reject'].shape[0]
            reject_val = group[group['Initial File Status (Credit)'] == 'Reject']['Request Amount'].sum() / 100000
            decision_count = sanction_count + reject_count
            decision_val = sanction_val + reject_val
            disbursed_count = group[group['Disb. Date'].notnull()].shape[0]
            disbursed_val = group[group['Disb. Date'].notnull()]['Sanction Amount'].sum() / 100000

            output_df1 = output_df1.append({
                'CCM': ccm,
                'No (Login)': login_count,
                'Val (In Lacs) (Login)': login_val,
                'No (Sanction)': sanction_count,
                'Val (In Lacs) (Sanction)': sanction_val,
                'No (Reject/Withdraw)': reject_count,
                'Val (In Lacs) (Reject/Withdraw)': reject_val,
                'No (Decision)': decision_count,
                'Val (In Lacs) (Decision)': decision_val,
                'No (Disbursed)': disbursed_count,
                'Val (In Lacs) (Disbursed)': disbursed_val
            }, ignore_index=True)

        totals1 = {
            'CCM': 'Total',
            'No (Login)': output_df1['No (Login)'].sum(),
            'Val (In Lacs) (Login)': output_df1['Val (In Lacs) (Login)'].sum(),
            'No (Sanction)': output_df1['No (Sanction)'].sum(),
            'Val (In Lacs) (Sanction)': output_df1['Val (In Lacs) (Sanction)'].sum(),
            'No (Reject/Withdraw)': output_df1['No (Reject/Withdraw)'].sum(),
            'Val (In Lacs) (Reject/Withdraw)': output_df1['Val (In Lacs) (Reject/Withdraw)'].sum(),
            'No (Decision)': output_df1['No (Decision)'].sum(),
            'Val (In Lacs) (Decision)': output_df1['Val (In Lacs) (Decision)'].sum(),
            'No (Disbursed)': output_df1['No (Disbursed)'].sum(),
            'Val (In Lacs) (Disbursed)': output_df1['Val (In Lacs) (Disbursed)'].sum()
        }

        output_df1 = output_df1.append(totals1, ignore_index=True)

        # Second DataFrame processing (from code2)
        output_df2 = pd.DataFrame(columns=[
            'CCM', 'No (Spill File)', 'Val (In Lacs) (Spill File)', 'No (Fresh Login)', 'Val (In Lacs) (Fresh Login)',
            'No (Total File)', 'Val (In Lacs) (Total File)', 'No (Sanction/Disbursed)', 'Val (In Lacs) (Sanction/Disbursed)',
            'No (Reject/Withdraw)', 'Val (In Lacs) (Reject/Withdraw)', 'No (Recommend)', 'Val (In Lacs) (Recommend)',
            'No (Query-Sales)', 'Val (In Lacs) (Query-Sales)', 'No (WIP-Credit)', 'Val (In Lacs) (WIP-Credit)',
            'No (Visit Pending)', 'Val (In Lacs) (Visit Pending)'
        ])

        user_date = selected_date
        current_month = user_date.strftime('%b')
        previous_month = (user_date.replace(day=1) - pd.DateOffset(months=1)).strftime('%b')

        for ccm, group in df.groupby(['CCM']):
            aug_group = group[group['MONTH'] == current_month]
            naug_group = group[group['MONTH'] == previous_month]

            spfno = len(naug_group)
            spfval = naug_group['Request Amount'].sum() / 100000
            frshlg = len(aug_group)
            frshval = aug_group['Request Amount'].sum() / 100000
            totalno = spfno + frshlg
            totalval = spfval + frshval
            sandis_count = group[group['Initial File Status (Credit)'] == 'Sanction'].shape[0]
            sandis_val = group[group['Initial File Status (Credit)'] == 'Sanction']['Sanction Amount'].sum() / 100000
            reject_count = group[group['Initial File Status (Credit)'] == 'Reject'].shape[0]
            reject_val = group[group['Initial File Status (Credit)'] == 'Reject']['Request Amount'].sum() / 100000
            rec_count = group[group['Initial File Status (Credit)'] == 'Recommend'].shape[0]
            rec_val = group[group['Initial File Status (Credit)'] == 'Recommend']['Request Amount'].sum() / 100000
            qs_count = group[group['Initial File Status (Credit)'] == 'Query- Sales'].shape[0]
            qs_val = group[group['Initial File Status (Credit)'] == 'Query- Sales']['Request Amount'].sum() / 100000
            wip_count = group[group['Initial File Status (Credit)'] == 'WIP- Credit'].shape[0]
            wip_val = group[group['Initial File Status (Credit)'] == 'WIP- Credit']['Request Amount'].sum() / 100000
            vp_count = group[group['Initial File Status (Credit)'] == 'Visit Pending'].shape[0]
            vp_val = group[group['Initial File Status (Credit)'] == 'Visit Pending']['Request Amount'].sum() / 100000

            output_df2 = output_df2.append({
                'CCM': ccm,
                'No (Spill File)': spfno,
                'Val (In Lacs) (Spill File)': spfval,
                'No (Fresh Login)': frshlg,
                'Val (In Lacs) (Fresh Login)': frshval,
                'No (Total File)': totalno,
                'Val (In Lacs) (Total File)': totalval,
                'No (Sanction/Disbursed)': sandis_count,
                'Val (In Lacs) (Sanction/Disbursed)': sandis_val,
                'No (Reject/Withdraw)': reject_count,
                'Val (In Lacs) (Reject/Withdraw)': reject_val,
                'No (Recommend)': rec_count,
                'Val (In Lacs) (Recommend)': rec_val,
                'No (Query-Sales)': qs_count,
                'Val (In Lacs) (Query-Sales)': qs_val,
                'No (WIP-Credit)': wip_count,
                'Val (In Lacs) (WIP-Credit)': wip_val,
                'No (Visit Pending)': vp_count,
                'Val (In Lacs) (Visit Pending)': vp_val
            }, ignore_index=True)

        # Save both dataframes to a temporary Excel file with two sheets
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
            output_df1.to_excel(writer, index=False, sheet_name='Report1')
            output_df2.to_excel(writer, index=False, sheet_name='Report2')

        temp_file_path = temp_file.name

        # Convert the first dataframe to an HTML table for preview
        output_html = output_df1.to_html(index=False)

        return send_file(temp_file_path, as_attachment=True, download_name='output.xlsx')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)




