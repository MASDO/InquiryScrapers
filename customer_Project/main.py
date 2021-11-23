# libraries
import matplotlib.pyplot as plt
# import squarify
# pip install squarify (algorithm for treemap)
import pandas as pd
import matplotlib
# Create a data frame with fake data
import plotly.express as px

if __name__ == "__main__":
    data = pd.read_csv("Customer_2.csv")
    # data['Effect']= data['Effect'].asfloat

    print(data)
#    df = pd.DataFrame({'Weights': [[0.45, 0.61], 0.15, 0.45, 0.12],
#                       'group': [["customer_trs", "Transactions"], "Card2C", "Deposits", "Deposits2"]})
#    print(df)
#    # plot it
#    squarify.plot(sizes=df['Weights'], label=df['group'], alpha=.6)
#    plt.axis('off')
#    plt.show()
#
# Create a dataset:
#    my_values = [i ** 3 for i in range(1, 100)]
#
#    # create a color palette, mapped to these values
#    cmap = matplotlib.cm.Blues
#    mini = min(my_values)
#    maxi = max(my_values)
#    norm = matplotlib.colors.Normalize(vmin=mini, vmax=maxi)
#    colors = [cmap(norm(value)) for value in my_values]
#
    # Change color
    fig = px.treemap(data,
                     path=['customer_type', 'Transaction_Type', 'Factor', 'Sub_Factor'], values='Effect')

    fig.update_layout(
         font_family="Tahoma",
         title_font_family="Tahoma"
    )
    fig.update_traces(legendgrouptitle_font_size=18)
    fig.show()


