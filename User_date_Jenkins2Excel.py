import argparse
import os
import pandas

def excel_update(usernm,pswd,Active_hrs,intervals,Activation):
    dict_new={}
    Excel_path = r"D:\BPL_SW\automate_vdi\jenkins_data.xls"
    os.makedirs(os.path.split(Excel_path)[0],exist_ok=True)
    if not os.path.exists(Excel_path):
        with open(Excel_path,'w') as fp:
            fp.close()
    if os.path.getsize(Excel_path):
        df=pandas.read_excel(Excel_path)
        # df = pandas.DataFrame.from_dict(df, orient='index')
        # dict1={'User-id' : [usernm],'Password':[pswd],'how many hrs to be active(in hrs)':[Active_hrs],'intervals(in mins)':[intervals],
        #            "last Updated" : [],"next_scheduled_time":[],"max_time":[],
        #             "Activation":[Activation]}
        # df1 = pandas.DataFrame.from_dict(dict1,orient='index')
        # df1=df1.transpose()
        # df.append(df1,ignore_index=True)
        # list1=[df,df1]
        # new_df=pandas.DataFrame(list1)
        # new_df.to_excel(Excel_path,index=False)
        # df.to_excel(Excel_path,index=False)

        for ind in df.index:
            dict_new['User-id'] =[df['User-id'][ind]]
            dict_new['Password']=[df['Password'][ind]]
            dict_new['how many hrs to be active(in hrs)']=[df['how many hrs to be active(in hrs)'][ind]]
            dict_new['intervals(in mins)']=[df['intervals(in mins)'][ind]]
            dict_new['last Updated']=[df['last Updated'][ind]]
            dict_new['next_scheduled_time']=[df['next_scheduled_time'][ind]]
            dict_new['max_time']=[df['max_time'][ind]]
            dict_new['Activation']=[df['Activation'][ind]]

        dict_new.setdefault('User-id',[]).append(usernm)
        dict_new.setdefault('Password',[]).append(pswd)
        dict_new.setdefault('how many hrs to be active(in hrs)',[]).append(Active_hrs)
        dict_new.setdefault('intervals(in mins)',[]).append(intervals)
        dict_new.setdefault('last Updated',[]).append(None)
        dict_new.setdefault('next_scheduled_time',[]).append(None)
        dict_new.setdefault('max_time',[]).append(None)
        dict_new.setdefault('Activation',[]).append(Activation)
        print(dict_new)

        df=pandas.DataFrame(dict_new)
        df.to_excel(Excel_path,index=False)

    else:
        dict1={'User-id' : [usernm],'Password':[pswd],'how many hrs to be active(in hrs)':[Active_hrs],'intervals(in mins)':[intervals],
                   "last Updated" : [],"next_scheduled_time":[],"max_time":[],
                    "Activation":[Activation]}
        df = pandas.DataFrame.from_dict(dict1, orient='index')
        df = df.transpose()
        df.to_excel(Excel_path, index=False)



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("username", type=str, help="windows userid")
    parser.add_argument("password", type=str, help="windows password")
    parser.add_argument("hrs_to_be_active_in_hrs", type=int)
    parser.add_argument("intervals_in_mins" , type=int)
    parser.add_argument("Activation", type=str, help="True/False")
    args = parser.parse_args()

    excel_update(args.username,args.password,args.hrs_to_be_active_in_hrs,args.intervals_in_mins,args.Activation)