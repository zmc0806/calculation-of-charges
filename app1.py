import streamlit as st
import pandas as pd
import io
import re
from pypinyin import lazy_pinyin  # 用于中文拼音排序
# 确保xlrd库已安装用于处理旧版.xls文件
# 如果运行时遇到错误，请安装：pip install xlrd>=2.0.1

st.set_page_config(page_title="月度业绩表分析", layout="wide")

st.title("月度业绩表分析")
st.sidebar.header("配置选项")

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
            
            # 处理单位数据
            st.subheader("单位业绩分析")
            
            # 创建单位业绩汇总
            try:
                # 提取相关列
                company_df = data[['单位名称', '收费方式', '金额', '师傅总路桥费', '代垫费', '外派金额']]
                
                # 按单位名称分组并汇总
                if '收费方式' in company_df.columns:
                    summary_df = (
                        company_df.groupby("单位名称")
                        .agg({
                            "收费方式": "first",  # 取第一个值
                            "金额": "sum",
                            "师傅总路桥费": "sum",
                            "代垫费": "sum",
                            "外派金额": "sum"
                        })
                        .reset_index()
                    )
                else:
                    # 如果没有收费方式列，创建默认值
                    summary_df = (
                        company_df.groupby("单位名称")
                        .agg({
                            "金额": "sum",
                            "师傅总路桥费": "sum",
                            "代垫费": "sum",
                            "外派金额": "sum"
                        })
                        .reset_index()
                    )
                    summary_df["收费方式"] = "签单"  # 添加默认收费方式
                
                # 按拼音排序
                summary_df["单位名称拼音"] = summary_df["单位名称"].apply(
                    lambda x: ''.join(lazy_pinyin(str(x)))
                )
                summary_df = summary_df.sort_values("单位名称拼音").drop(columns="单位名称拼音")
                
                # 筛选掉单位名称中含有"先生"或"女士"的行
                summary_df = summary_df[~summary_df["单位名称"].str.contains("先生|女士", na=False)]
                
                # 在"收费方式"右边添加一个"地区"列
                col_index = summary_df.columns.get_loc("收费方式") + 1
                summary_df.insert(col_index, "地区", "")  # 空字符串填充
                
                # 在"金额"右边添加"开票金额"和"到账时间"列
                col_index = summary_df.columns.get_loc("金额") + 1
                summary_df.insert(col_index, "开票金额", "")
                summary_df.insert(col_index + 1, "到账时间", "")
                
                # 显示单位业绩汇总表
                st.dataframe(summary_df)
                
                # 允许用户编辑地区信息
                st.subheader("添加地区信息")
                
                # 获取所有唯一单位名称
                all_companies = summary_df["单位名称"].unique().tolist()
                
                # 创建列以显示地区编辑界面
                col1, col2 = st.columns(2)
                
                with col1:
                    # 创建一个字典来存储每个单位的地区
                    regions = {}
                    unique_regions = ["成都市", "绵阳市", "德阳市", "自贡市", "泸州市", "其他"]
                    
                    # 添加选项让用户添加新的地区
                    new_region = st.text_input("添加新地区")
                    if new_region and new_region not in unique_regions:
                        unique_regions.append(new_region)
                    
                    selected_region = st.selectbox("选择地区", unique_regions)
                    selected_companies = st.multiselect("选择单位", all_companies)
                    
                    if st.button("分配地区"):
                        for company in selected_companies:
                            regions[company] = selected_region
                        st.success(f"已将 {len(selected_companies)} 个单位分配到 {selected_region}")
                
                # 创建收款表
                st.subheader("收款汇总表（按地区）")
                
                # 将当前选择的地区分配应用到summary_df
                for company, region in regions.items():
                    summary_df.loc[summary_df["单位名称"] == company, "地区"] = region
                
                # 创建收款汇总表（按地区分组）
                if regions:  # 只有当有地区分配时才创建
                    # 过滤掉没有地区的行
                    region_df = summary_df[summary_df["地区"] != ""].copy()
                    
                    if not region_df.empty:
                        # 按地区分组并汇总金额
                        region_summary = region_df.groupby("地区").agg({
                            "金额": "sum",
                            "开票金额": "sum",  # 这将合计非空值
                            "师傅总路桥费": "sum",
                            "代垫费": "sum",
                            "外派金额": "sum"
                        }).reset_index()
                        
                        # 显示收款汇总表
                        st.dataframe(region_summary)
                    else:
                        st.info("尚未分配地区信息，无法生成收款汇总表")
                else:
                    st.info("请先分配地区信息以生成收款汇总表")
                
                # Data visualization
                st.subheader("数据可视化")
                
                tab1, tab2 = st.tabs(["师傅业绩", "单位业绩"])
                
                with tab1:
                    # Bar chart for total amount by master
                    st.subheader("师傅金额分布")
                    st.bar_chart(sorted_df.set_index('师傅姓名')['金额'])
                    
                    # Bar chart for total orders by master
                    st.subheader("师傅单数分布")
                    st.bar_chart(sorted_df.set_index('师傅姓名')['总单数'])
                
                with tab2:
                    # Bar chart for total amount by company
                    st.subheader("单位金额分布")
                    st.bar_chart(summary_df.set_index('单位名称')['金额'])
                    
                    # 如果有地区数据，显示地区分布
                    if regions:
                        region_data = summary_df[summary_df["地区"] != ""]
                        if not region_data.empty:
                            st.subheader("地区金额分布")
                            st.bar_chart(region_data.groupby("地区").sum()['金额'])
                
                # Download the processed data
                st.subheader("数据导出")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    sorted_df.to_excel(writer, sheet_name=f'{month}月师傅业绩', index=False)
                    summary_df.to_excel(writer, sheet_name=f'{month}月单位业绩', index=False)
                    
                    # 如果有地区数据，添加收款汇总表
                    if regions and not region_df.empty:
                        region_summary.to_excel(writer, sheet_name=f'{month}月收款汇总', index=False)
                
                output.seek(0)
                
                # Use the same file extension as the uploaded file
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_ext == "xlsx" else "application/vnd.ms-excel"
                
                st.download_button(
                    label="下载处理后的Excel文件",
                    data=output,
                    file_name=f"{month}月业绩汇总.{file_ext}",
                    mime=mime_type
                )
            except Exception as e:
                st.error(f"处理单位数据时出错: {e}")
            
        else:
            st.error("文件格式不符合要求，请检查是否包含必要的列：师傅姓名、金额、师傅总路桥费、代垫费")
    
    except Exception as e:
        st.error(f"处理文件时出错: {e}")
else:
    st.info("请上传月度业绩表Excel文件")

# Add some information at the bottom
st.markdown("---")
st.markdown("说明：此应用程序将分析月度业绩表，计算每位师傅的总金额、总单数、总路桥费和代垫费。同时提供单位业绩汇总以及按地区的收款汇总表。")
st.markdown("上传文件格式应为'X月业绩表.xlsx'或'X月业绩表.xls'，其中X为1-12。")
st.markdown("注意：为使用拼音排序功能，请确保安装了pypinyin库：`pip install pypinyin`")
