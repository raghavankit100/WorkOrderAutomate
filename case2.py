 if row[3] is not None:   #(IF Group_Ref is not null)

                        # Gourp Id is not none so we have to check they all are same or not
                        same_value = all(row2[3] == client_rows[0][3] for row2 in client_rows)
                        if same_value:
                            print(f'{row[3]} Group_ref has been processed')
                            desired_date_format = row[4] # Extract only the date part (DD-MM-YYYY) we have to do this in query only particular format is working
                            # Step 2: Construct the query with TO_DATE using the desired format
                            GrpRef_query = f"""SELECT UMR, Created_Date, Workorder_Ref, NO_IN_GROUP, WORKORDER_TAG, BROKER_CODE, GROUP_REFERENCE, WORK_ORDER_STATUS FROM repository.tblworkorder WHERE Group_Reference = '{row[3]}' AND created_date = TO_DATE('{desired_date_format}', 'DD-MM-YYYY')"""
                            
                            table=[]
                            for i in cur.execute(GrpRef_query):
                                table.append(i)   
                            # Now table have all the rows that have above query data
                            if any(status[7] != 'CAN' for status in table):
                                print("There is at least one work order with a status other than 'NEW' while checking for case2")                    
                            # check status, we already did but still for confirmation
                            else:
                                def check_column(table):
                                    reference_value = table[0][5]
                                    different_value = None
                                    
                                    for row in table:
                                        if row[5] != reference_value:
                                            different_value = row[5]
                                            break
                                            
                                    return reference_value, different_value
                                
                                reference_value, different_value = check_column(table)
                                #this will store same brokercode as well as different broker code if there is any
                                
                                #CASE 2 
                                noInGroup=table[0][3]
                                submissions=len(table)
                                if different_value is None: #Broker Code is unique 
                                    
                                    if(submissions != noInGroup):
                                        print(f'submission are {submissions} whereas noInGroup are {noInGroup} \n')
                                        recipient_email = "ankit.raghav@dxc.com"
                                        text_content=f"""Hi, <br><br>Below is the Work package that you have submitted on {EmailDATE}, which has failed to reach some of the required systems at our end.<br><br>
                                        The details are as follows"""  
                                        
                                        
                                        text_content2=f"""Please not that these submissions has been failed because you have entered no in group as {noInGroup} whereas {submissions} has been made.<br><br>
                                        Hence, we would advise you make these submission again and we will cancel the work packages at our end. If you have any further concerns please raise a call with the Xchanging Service Centre.<br><br><br>
                                        Thanks and Regards<br>
                                        Ankit<br>
                                        DXC Technology<br>
                                        """
                                        html_content="""<!DOCTYPE html>
                                        <html lang="en">
                                        <head>
                                        <meta charset="UTF-8">
                                        <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                        <title>Table Example</title>
                                        <style>
                                            table {
                                                border-collapse: collapse;
                                                width: 50%;
                                            }
                                            th, td {
                                                border: 1px solid black;
                                                padding: 8px;
                                                text-align: center;
                                            }
                                            th {
                                                background-color: #f2f2f2;
                                            }
                                            td[colspan="3"] {
                                                background-color: #c0c0c0;
                                                font-weight: bold;
                                            }
                                        </style>
                                        </head>
                                        <body>
                                        <table>
                                        <thead>
                                            <tr>
                                            <th>UMR</th>
                                            <th>WORKORDER_TAG</th>
                                            <th>WORKORDER_REF</th>
                                            <th>Group_REF</th>
                                            <th>No_in_Group</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            
                                            {% for row in client_rows%}
                                            <tr>
                                            <td>{{ row[0] }}</td>
                                            <td>{{ row[1] }}</td>
                                            <td>{{ row[22] }}</td>
                                            <td>{{ row[3] }}</td>
                                            <td>{{ noInGroup }}</td>
                                            </tr>
                                            {% endfor %}  
                                            
                                        </tbody>
                                        </table>
                                        </body>
                                        </html>
                                        """
                                        template = Template(html_content)
                                        # Render the template with client_rows data
                                        rendered_html = template.render(client_rows=client_rows,noInGroup=noInGroup)
                                        send_email(text_content,rendered_html,text_content2,recipient_email)
                                #Case3 if Broker code have different value
                                elif different_value is not None:
                                    print(f"It have d/f broker code {different_value}")
                                    recipient_email = "ankit.raghav@dxc.com"
                                    text_content=f"""Hi, <br><br>Below is the Work package that you have submitted on {row[16]}, which has failed to reach some of the required systems at our end.<br><br>
                                    The details are as follows"""  
                                    text_content2="""Please note that these submissions have failed because you have made submission with different broker codes.<br><br>
                                    Hence, we would advise you make these submission again with same broker code and with correct number in group. The current work packages will be cancelled by the system at our end. If you have any further concerns please raise a call with the Xchanging Service Centre.<br><br><br>
                                    Thanks and Regards<br>
                                    Ankit<br>
                                    DXC Technology<br>
                                    """
                                    html_content="""<!DOCTYPE html>
                                    <html lang="en">
                                    <head>
                                    <meta charset="UTF-8">
                                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                    <title>Table Example</title>
                                    <style>
                                        table {
                                            border-collapse: collapse;
                                            width: 50%;
                                        }
                                        th, td {
                                            border: 1px solid black;
                                            padding: 8px;
                                            text-align: center;
                                        }
                                        th {
                                            background-color: #f2f2f2;
                                        }
                                        td[colspan="3"] {
                                            background-color: #c0c0c0;
                                            font-weight: bold;
                                        }
                                    </style>
                                    </head>
                                    <body>
                                    <table>
                                    <thead>
                                        <tr>
                                        <th>UMR</th>
                                        <th>WORKORDER_TAG</th>
                                        <th>WORKORDER_REF</th>
                                        <th>Group Reference</th>
                                        <th>No_In_Group</th>
                                        <th>Broker Number</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {% set is_first_row = True %}
                                        {% for row in client_rows %}
                                        <tr>
                                        <td>{{ row[0] }}</td>
                                        <td>{{ row[1] }}</td>
                                        <td>{{ row[22] }}</td>
                                        {% if is_first_row %}
                                        <td>{{ row[3]  }}</td>
                                        <td>{{ noInGroup }}</td>
                                        <td>{{ reference_value,different_value }}</td>
                                        {% else %}
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        {% endif %}
                                        {% set is_first_row = False %}
                                        </tr>
                                        {% endfor %}
                                    </tbody>
                                    </table>
                                    </body>
                                    </html>
                                    """
                                    template = Template(html_content)
                                    rendered_html = template.render(client_rows=client_rows,noInGroup=noInGroup,reference_value=reference_value, different_value=different_value)
                                    send_email(text_content,rendered_html,text_content2,recipient_email)
                                #Wog case still in process need to complete
                                else:
                                    wog_query= f'SELECT * FROM repository.tblworkordergroup WHERE wog_group_reference="{row[3]}"'
                                    pass
                                    #Check the query and update the code with new steps and conditions
                            #to update the excel from missing to tracker+workflowAQ
                            updateExcel(row)
                        else:
                            print("Some rows have different GroupRef")
                        break