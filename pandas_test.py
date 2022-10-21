import pandas as pd

cust_name = "Vertikal Tech"
master_list = pd.read_excel(f'{cust_name} TC Plant Inventory List.xlsx', sheet_name = "Master List")

def create_stage1_list():
    stage1_ready = master_list.loc[master_list["Stage I Ready"] == "Yes"]
    stage1_df = str(stage1_ready[['Variety', "Stage I Date"]].sort_values("Stage I Date"))
    with open(f"{cust_name} stage1_list", "w") as stage1_list:
        stage1_list.write(f'{cust_name}\n')
        stage1_list.write(stage1_df)

def create_stage3_and_MS_list():
    stage3_ready = master_list.loc[master_list["Stage III Ready"] == "Yes"]
    stageMS_ready = master_list.loc[master_list["Stage 2.5 Ready"] == "Yes"]
    stage3_df = str(stage3_ready[['Variety', "Stage III Date"]].sort_values("Stage III Date"))
    stageMS_df = str(stageMS_ready[['Variety', "Stage 2.5 Date"]].sort_values("Stage 2.5 Date"))
    with open(f"{cust_name} stage3_list", "w") as stage3_list:
        stage3_list.write(f'{cust_name}\n')
        stage3_list.write(stage3_df)
        stage3_list.write("\n\n")
        stage3_list.write(stageMS_df)

def main():
    create_stage1_list()
    create_stage3_and_MS_list()

if __name__ == "__main__":
    main()
