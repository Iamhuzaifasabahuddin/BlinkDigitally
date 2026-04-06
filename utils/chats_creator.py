import plotly.express as px
import plotly.graph_objects as go

def create_review_pie_chart(review_data: dict[str, int], title: str):
    """Create pie chart for review distribution"""
    global labels, values
    if isinstance(review_data, dict):
        if not review_data or sum(review_data.values()) == 0:
            return None
        values = list(review_data.values())
        labels = list(review_data.keys())

    custom_colors = {
        "Attained": "#7dff8d",
        "Pending": "#ffc444",
        "Negative": "#ff4b4b",
        "Sent": "#77e5f7"
    }

    fig = px.pie(
        values=values,
        names=labels,
        title=title,
        color=labels,
        color_discrete_map=custom_colors
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return fig


def create_platform_comparison_chart(usa_data: dict[str, int], uk_data: dict[str, int]):
    """Create comparison chart for platforms"""
    platforms = ['Amazon', 'Barnes & Noble', 'Ingram Spark', "Draft2Digital", "Kobo", "LULU", "FAV", "ACX"]

    fig = go.Figure(data=[
        go.Bar(name='USA', x=platforms, y=list(usa_data.values()), marker_color="#23A0F8"),
        go.Bar(name='UK', x=platforms, y=list(uk_data.values()), marker_color="#ff7f0e")
    ])

    fig.update_layout(
        title='Platform Distribution: USA vs UK',
        barmode='group',
        xaxis_title='Platforms',
        yaxis_title='Number of Reviews'
    )
    return fig


def create_brand_chart(usa_brands: dict[str, int], uk_brands: dict[str, int]):
    """Create brand distribution chart"""
    all_brands = list(usa_brands.keys()) + list(uk_brands.keys())
    all_values = list(usa_brands.values()) + list(uk_brands.values())
    regions = ['USA'] * len(usa_brands) + ['UK'] * len(uk_brands)

    fig = px.bar(
        x=all_brands,
        y=all_values,
        color=regions,
        title='Brand Distribution by Region',
        color_discrete_map={'USA': '#23A0F8', 'UK': '#ff7f0e'},
        labels={
            "x": "Brands",
            "y": "Number of Clients"
        }
    )
    return fig