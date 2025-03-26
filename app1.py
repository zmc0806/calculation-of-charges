import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="月度业绩表分析", layout="wide")

st.title("月度业绩表分析")

uploaded_file = st.file_uploader("上传业绩表 (格式: X月业绩表.xlsx 或 X月业绩表.xls，X为1-12)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Extract month from filename
    filename = uploaded_file.name
    month_match = re.search(r'(\d+)月业绩表', filename)
    month = month_match.group(1) if month_match else "未知"
    
    # Determine file extension
    file_ext = "xlsx" if filename.lower().endswith("xlsx") else "xls"
    
    st.subheader(f"{month}月业绩表分析")
    
    # Read the Excel file
    try:
        data = pd.read_excel(uploaded_file)
        
        # Display the raw data
        with st.expander("查看原始数据"):
            st.dataframe(data)
        
        # Clean and process the data
        if '师傅姓名' in data.columns and '金额' in data.columns:
            # Extract relevant columns
            clean = data[['师傅姓名', '金额', '师傅总路桥费', '代垫费']]
            
            # Sum by 师傅姓名
            qw = clean.groupby('师傅姓名').sum()
            qw = qw.reset_index()
            
            # Count records by 师傅姓名
            re = data[['师傅姓名', '流水号']]
            re = re.rename(columns={'流水号': '总单数'})
            re = re.groupby('师傅姓名').count()
            
            # Merge the dataframes
            merged_df = pd.merge(qw, re, on='师傅姓名')
            
            # Define consistent order for 师傅姓名 (can be customized)
            default_order = ['蔡勇', '陈行辉', '高勇军', '孙琪琪', '孙涛', '唐正荣', '萧敏', '杨彬', '姚强']
            
            # Allow users to customize the order
            st.subheader("自定义师傅排序顺序")
            custom_order = st.text_area("输入师傅姓名顺序（每行一个名字）", 
                                    value="\n".join(default_order))
            
            order = [name.strip() for name in custom_order.split("\n") if name.strip()]
            
            if order:
                # Filter the DataFrame to include only masters in the order list and others
                in_order = merged_df[merged_df['师傅姓名'].isin(order)]
                not_in_order = merged_df[~merged_df['师傅姓名'].isin(order)]
                
                # Set category and order for those in the order list
                in_order['师傅姓名'] = pd.Categorical(in_order['师傅姓名'], 
                                            categories=order, 
                                            ordered=True)
                
                # Sort by the categorical order
                sorted_in_order = in_order.sort_values(by='师傅姓名')
                
                # Combine sorted and unsorted parts
                sorted_df = pd.concat([sorted_in_order, not_in_order])
            else:
                sorted_df = merged_df
            
            # Display the processed data
            st.subheader("处理后的数据")
            st.dataframe(sorted_df)
            
            # Calculate summary statistics
            total_amount = sorted_df['金额'].sum()
            total_orders = sorted_df['总单数'].sum()
            total_tolls = sorted_df['师傅总路桥费'].sum()
            total_advances = sorted_df['代垫费'].sum()
            
            # Display summary statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("总金额", f"{total_amount:,.2f}")
            with col2:
                st.metric("总单数", f"{total_orders}")
            with col3:
                st.metric("总路桥费", f"{total_tolls:,.2f}")
            with col4:
                st.metric("总代垫费", f"{total_advances:,.2f}")
            
            # Data visualization
            st.subheader("数据可视化")
            
            # Bar chart for total amount by master
            st.bar_chart(sorted_df.set_index('师傅姓名')['金额'])
            
            # Bar chart for total orders by master
            st.bar_chart(sorted_df.set_index('师傅姓名')['总单数'])
            
            # Download the processed data
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sorted_df.to_excel(writer, sheet_name=f'{month}月新', index=False)
            
            output.seek(0)
            
            # Use the same file extension as the uploaded file
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_ext == "xlsx" else "application/vnd.ms-excel"
            
            st.download_button(
                label="下载处理后的Excel文件",
                data=output,
                file_name=f"{month}月新.{file_ext}",
                mime=mime_type
            )
            
        else:
            st.error("文件格式不符合要求，请检查是否包含必要的列：师傅姓名、金额、师傅总路桥费、代垫费")
    
    except Exception as e:
        st.error(f"处理文件时出错: {e}")
else:
    st.info("请上传月度业绩表Excel文件")

# Add some information at the bottom
st.markdown("---")
st.markdown("说明：此应用程序将分析月度业绩表，计算每位师傅的总金额、总单数、总路桥费和代垫费。")
st.markdown("上传文件格式应为'X月业绩表.xlsx'或'X月业绩表.xls'，其中X为1-12。")