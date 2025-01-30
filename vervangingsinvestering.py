import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta

# Set page config
st.set_page_config(page_title="Asset Management Dashboard", layout="wide")

def format_euro(value):
    """Format a number as Euro currency with Dutch formatting (. for thousands, , for decimals)"""
    if pd.isna(value) or value == 0:
        return "€ 0"
    # Format with international notation first (1234.56)
    formatted = f"{float(value):,.2f}"
    # Convert to Dutch notation (1.234,56)
    formatted = formatted.replace(",", "@").replace(".", ",").replace("@", ".")
    return f"€ {formatted}"

def clean_number(value):
    try:
        if isinstance(value, str):
            cleaned = value.replace(',', '').replace('.', '')
            if cleaned.strip():
                try:
                    return float(cleaned)
                except ValueError:
                    st.error(f"Could not convert value to number: {value}")
                    return 0.0
        elif isinstance(value, (int, float)):
            return float(value)
        return 0.0
    except Exception as e:
        st.error(f"Error in clean_number: {str(e)}, value: {value}, type: {type(value)}")
        return 0.0

# Read the Excel file
@st.cache_data
def load_data():
    try:
        excel_path = Path("Archive/copied_values.xlsx")
        if not excel_path.exists():
            st.error(f"Excel file not found at: {excel_path}")
            return pd.DataFrame(columns=['Object', 'Waarde'])
        
        df = pd.read_excel(excel_path)
        df = pd.DataFrame({
            'Object': df.iloc[:, 0].astype(str),
            'Waarde': df.iloc[:, 2]
        })
        
        df = df[:-1].copy()
        df['Waarde'] = df['Waarde'].apply(clean_number)
        df['Object'] = df['Object'].str.strip()
        
        return df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        return pd.DataFrame(columns=['Object', 'Waarde'])

def create_section(title, df, factor_range=(0.8, 1.2), show_factors=False):
    try:
        st.header(title)
        
        section_df = df.copy()
        
        if section_df.empty:
            st.error("No data available")
            return None
            
        # Initialize factors in session state if not exists
        if f'factors_{title}' not in st.session_state:
            np.random.seed(42 if title == "Vervangingsinvesteringen" else 43)
            factors_df = pd.DataFrame({
                'Object': section_df['Object'],
                'Factor': [np.random.uniform(factor_range[0], factor_range[1]) for _ in range(len(section_df))]
            })
            st.session_state[f'factors_{title}'] = factors_df
        
        # Display editable factors table only for Exploitatiebudget
        if show_factors:
            st.subheader("Vermenigvuldigingsfactoren")
            factors_df = st.data_editor(
                st.session_state[f'factors_{title}'],
                column_config={
                    "Object": "Object",
                    "Factor": st.column_config.NumberColumn(
                        "Factor",
                        min_value=0.0,
                        max_value=2.0,
                        step=0.01,
                        format="%.2f"
                    )
                },
                hide_index=True,
                key=f"editor_{title}"
            )
            st.session_state[f'factors_{title}'] = factors_df
        else:
            factors_df = st.session_state[f'factors_{title}']
        
        # Create factors dictionary
        factors_dict = dict(zip(factors_df['Object'], factors_df['Factor']))
        
        # Apply factors and calculate results
        section_df['Vermenigvuldigingsfactor'] = section_df['Object'].map(factors_dict)
        section_df['Resultaat'] = section_df['Waarde'] * section_df['Vermenigvuldigingsfactor']
        
        # Create layout
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            st.subheader("Gegevenstabel")
            display_df = pd.DataFrame({
                'Object': section_df['Object'],
                'Oorspronkelijke Waarde': section_df['Waarde'].apply(format_euro),
                'Factor': section_df['Vermenigvuldigingsfactor'].round(3),
                'Resultaat': section_df['Resultaat'].apply(format_euro)
            })
            st.dataframe(display_df, hide_index=True)
        
        with col2:
            st.subheader("Vergelijking Waarden")
            fig_bar = px.bar(section_df, x='Object', y='Waarde',
                           title="Verdeling per Object")
            fig_bar.update_layout(
                height=400,
                margin=dict(t=30, b=0),
                xaxis_tickangle=45
            )
            st.plotly_chart(fig_bar, use_container_width=True, key=f"bar_chart_{title}")
        
        with col3:
            st.subheader("Verdeling")
            fig_pie = px.pie(section_df, values='Waarde', names='Object',
                           title="Verdeling per Object")
            fig_pie.update_layout(
                height=400,
                margin=dict(t=30, b=0)
            )
            st.plotly_chart(fig_pie, use_container_width=True, key=f"pie_chart_{title}")
        
        # Display totals
        total_col1, total_col2 = st.columns(2)
        with total_col1:
            original_total = section_df['Waarde'].sum()
            st.metric(
                "Totale Oorspronkelijke Waarde",
                format_euro(original_total)
            )
        with total_col2:
            new_total = section_df['Resultaat'].sum()
            difference = new_total - original_total
            st.metric(
                "Totale Waarde Na Vermenigvuldiging",
                format_euro(new_total),
                delta=format_euro(difference)
            )
        
        return section_df
    except Exception as e:
        st.error(f"Error in create_section: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        return None

try:
    # Load the data
    df = load_data()
    
    # Create Vervangingsinvesteringen section (without factors table)
    df_verv = create_section("Vervangingsinvesteringen", df, (0.8, 1.2), show_factors=False)
    
    st.markdown("---")
    
    # Create Exploitatiebudget section with factors table
    df_expl = create_section("Exploitatiebudget", df, (0.9, 1.3), show_factors=True)
    
    # Timeline visualization
    st.markdown("---")
    st.header("Afwaardering Tijdlijn")
    
    # Initialize timeline data in session state if not exists
    if 'timeline_data' not in st.session_state:
        np.random.seed(44)
        timeline_df = pd.DataFrame({
            'Object': df['Object'],
            'Jaren': [np.random.randint(10, 51) for _ in range(len(df))]
        })
        st.session_state.timeline_data = timeline_df
    
    # Display editable timeline table
    timeline_df = st.data_editor(
        st.session_state.timeline_data,
        column_config={
            "Object": "Object",
            "Jaren": st.column_config.NumberColumn(
                "Afwaardering Periode (jaren)",
                min_value=10,
                max_value=50,
                step=1,
                format="%d"
            )
        },
        hide_index=True,
        key="timeline_editor"
    )
    st.session_state.timeline_data = timeline_df
    
    # Create timeline visualization
    current_year = datetime.now().year
    
    # Create a figure for the timeline
    fig_timeline = go.Figure()
    
    # Add bars for each object
    for _, row in timeline_df.iterrows():
        fig_timeline.add_trace(go.Bar(
            name=row['Object'],
            x=[int(row['Jaren'])],  # Convert to int to ensure proper display
            y=[row['Object']],
            orientation='h',
            text=f"{int(row['Jaren'])} jaar",
            textposition='auto',
            width=0.7,  # Adjust bar width
            hovertemplate=(
                f"Object: {row['Object']}<br>" +
                f"Afwaardering: {int(row['Jaren'])} jaar<br>" +
                f"Van: {current_year}<br>" +
                f"Tot: {current_year + int(row['Jaren'])}"
            )
        ))
    
    # Update layout
    fig_timeline.update_layout(
        title="Afwaardering Periode per Object",
        xaxis_title="Jaren",
        yaxis_title="Object",
        showlegend=False,
        height=600,
        barmode='overlay',
        xaxis=dict(
            range=[0, 55],  # Set range slightly larger than max years
            tickmode='linear',
            tick0=0,
            dtick=5
        ),
        margin=dict(l=200, r=20, t=30, b=50)  # Adjust margins to show full object names
    )
    
    st.plotly_chart(fig_timeline, use_container_width=True, key="timeline_chart")
    
    # Add summary table for timeline
    st.subheader("Afwaardering Details")
    timeline_display = pd.DataFrame({
        'Object': timeline_df['Object'],
        'Waarde': [format_euro(df.loc[df['Object'] == obj, 'Waarde'].iloc[0]) for obj in timeline_df['Object']],
        'Afwaardering Periode': timeline_df['Jaren'].apply(lambda x: f"{int(x)} jaar"),
        'Start Jaar': current_year,
        'Eind Jaar': timeline_df['Jaren'].apply(lambda x: current_year + int(x))
    })
    st.dataframe(timeline_display, hide_index=True)

except Exception as e:
    st.error(f"Er is een fout opgetreden: {str(e)}")
    st.error("Controleer of het Excel bestand aanwezig is en de juiste structuur heeft.")
