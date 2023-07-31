def flooresistance_combined_graph(df):
    plt.figure(figsize=(16, 8))

    # bar graph
    plt.subplot(121)  # Sort the DataFrame by "Location" in ascending order
    x = df["Location"]
    y = df["Effective Insulation Resistance (kΩ)"]
    colors = ["#b967ff", "#e0a899", "#fffb96", "#428bca"]  # Add more colors if needed
    sorted_indices = np.argsort(y)  # Sort the indices based on y values
    x_sorted = [x[i] for i in sorted_indices]
    y_sorted = [y[i] for i in sorted_indices]
    bars = plt.bar(x_sorted, y_sorted, color=colors)
    plt.xlabel("Location")
    plt.ylabel("Effective Insulation Resistance (kΩ)")
    plt.title("Location VS Effective Floor Resistance Bar Graph")
    for bar, value in zip(bars, y_sorted):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    df["Result"] = flooresistance_rang(df.shape[0])
    df_counts = df["Result"].value_counts()
    labels = df_counts.index.tolist()
    values = df_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.title("Test Results")
    plt.axis("equal")
    # Save the combined graph as bytes
    graph_combined1 = io.BytesIO()
    plt.savefig(graph_combined1)
    plt.close()

    return graph_combined1


def insulation_combined_graph(mf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    x = mf["Location"]
    y = mf["Insulation Resistance (MΩ)"]
    colors = ["#b967ff", "#e0a899", "#fffb96", "#428bca"]  # Add more colors if needed
    sorted_indices = np.argsort(y)  # Sort the indices based on y values
    x_sorted = [x[i] for i in sorted_indices]
    y_sorted = [y[i] for i in sorted_indices]
    bars = plt.bar(x_sorted, y_sorted, color=colors)
    plt.xlabel("Location")
    plt.ylabel("Insulation Resistance (MΩ)")
    plt.title("Location by Insulation Resistance (MΩ)")
    plt.yscale("log")
    for bar, value in zip(bars, y_sorted):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    mf["Result"] = insulation_rang(mf.shape[0])
    mf_counts = mf["Result"].value_counts()
    labels = mf_counts.index.tolist()
    values = mf_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.title("Test Results")
    plt.axis("equal")

    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def phase_combined_graph(pf):
    plt.figure(figsize=(16, 8))
    # Bar graph
    plt.subplot(121)
    x = pf["Phase Sequence"]
    y = (
        pf["Phase Sequence"].value_counts().values
    )  # Use value_counts().values to get the count values as an array
    colors = ["#FFFF00", "#0000FF"]  # Add more colors if needed
    bars = plt.bar(x, y, color=colors)
    plt.ylabel("Count of Phase Sequence")
    plt.xlabel("Phase sequence")
    plt.title(" Phase Sequence Test")
    plt.yticks(np.arange(0, max(y) + 1, 1))
    for bar, value in zip(bars, y):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    pf["Result"] = phase_rang(pf)  # Ensure you have the phase_rang() function defined correctly
    pf_counts = pf["Result"].value_counts()
    labels = pf_counts.index.tolist()
    values = pf_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.axis("equal")
    plt.title("Test Results")
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def polarity_combined_graph(af):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)  # Prepare the data for the graph
    af["Result"] = polarity_rang(
        af.shape[0]
    )  # Ensure you have the polarity_rang() function defined correctly
    result_counts = af["Result"].value_counts().sort_index()
    x = result_counts.index
    y = result_counts.values
    colors = ["#b967ff", "#e0a899", "#fffb96", "#428bca"]
    bars = plt.bar(x, y, color=colors)
    plt.xlabel("Result")
    plt.ylabel("Count")
    plt.title("Result Counts")  # Setting y-axis ticks to integers
    plt.yticks(np.arange(y.max() + 1))
    for bar, value in zip(bars, y):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    af["Result"] = polarity_rang(
        af.shape[0]
    )  # Ensure you have the polarityrang() function defined correctly
    af_counts = af["Result"].value_counts()
    labels = af_counts.index.tolist()
    values = af_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.axis("equal")
    plt.title("Polarity Results")
    # Save the combined graph as bytes
    graph_combined4 = io.BytesIO()
    plt.savefig(graph_combined4)
    plt.close()

    return graph_combined4


def voltage_combined_graph(vf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    x = vf["Conductor Type"] + ", " + vf["Cable Length (m)"].astype(str)
    y = vf["Voltage Drop %"]
    colors = ["#d9534f", "#5bc0de", "#5cb85c", "#428bca"]
    sorted_indices = np.argsort(y)  # Sort the indices based on y values
    x_sorted = [x[i] for i in sorted_indices]
    y_sorted = [y[i] for i in sorted_indices]
    bars = plt.bar(x_sorted, y_sorted, color=colors)
    plt.ylabel("Voltage Drop %")
    plt.xlabel("Conductor type and Cable Length (m)")
    plt.title("Conductor type and Cable Length (m) VS Voltage Drop %")

    # Add values inside each bar vertically
    for bar, value in zip(bars, y_sorted):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    vf["Result"] = voltage_rang(
        len(vf)
    )  # Ensure you have the voltage_rang() function defined correctly
    vf_counts = vf["Result"].value_counts()
    labels = vf_counts.index.tolist()
    values = vf_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.axis("equal")
    plt.title("Test Results")
    # Save the combined graph as bytes
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def residual_combined_graph(rf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    tct = rf["Trip curve type"].value_counts()
    x_values = tct.index
    y_values = tct.values
    colors = ["#b967ff", "#e0a899"]  # Add more colors if needed
    bars = plt.bar(x_values, y_values, color=colors)
    plt.xlabel("Trip curve type")
    plt.ylabel("Count")
    plt.title("Residual Test Results")
    for bar, value in zip(bars, y_values):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    # rf['Result'] = residual_rang(rf.shape[0])  # Ensure you have the residual_rang() function defined correctly
    result_counts = rf["Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", shadow=False, startangle=90, colors=colors)
    plt.title("Residual Test Results")
    plt.axis("equal")  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def Earth_combined_graph(ef):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    x = ef["Location"]
    y = ef["Measured Earth Resistance - Individual"]
    colors = ["#b967ff", "#e0a899", "#fffb96", "#428bca"]  # Add more colors if needed
    sorted_indices = np.argsort(y)  # Sort the indices based on y values
    x_sorted = [x[i] for i in sorted_indices]
    y_sorted = [y[i] for i in sorted_indices]
    bars = plt.bar(x_sorted, y_sorted, color=colors)
    plt.xlabel("Location")
    plt.ylabel("Measured Earth Resistance - Individual")
    plt.title("Location VS Measured Earth Resistance - Individual graph")
    for bar, value in zip(bars, y_sorted):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # pie diagram
    plt.subplot(122)
    result_counts = ef["Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", shadow=False, startangle=90, colors=colors)
    plt.title("Earth Pit Test Results")
    plt.axis("equal")  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()
    return graph_combined


def threephase_combined_graph(tf):
    # Create a 2x2 grid of subplots
    fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(16, 8))

    # First subplot - Bar graph for Voltage Unbalance %
    ax1 = axes[0, 0]
    x = tf["Location "]
    y = tf["Voltage Unbalance %"]
    colors = ["#5cb85c", "#428bca"]
    bars = ax1.bar(x, y, color=colors)
    ax1.set_xlabel("Location")
    ax1.set_ylabel("Voltage Unbalance %")
    ax1.set_title("Location VS Voltage Unbalance %")
    for bar, value in zip(bars, y):
        height = bar.get_height()
        ax1.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}%",
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Second subplot - Pie chart for Current Unbalance %
    ax2 = axes[0, 1]
    x = tf["Location "]
    y = tf["Current Unbalance %"]
    colors = ["#d9534f", "#5bc0de"]
    bars = ax2.bar(x, y, color=colors)
    ax2.set_xlabel("Location")
    ax2.set_ylabel("Current Unbalance %")
    ax2.set_title("Location VS Current Unbalance %")
    for bar, value in zip(bars, y):
        height = bar.get_height()
        ax2.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}%",
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Third subplot - Bar graph for Voltage-NE (V)
    ax3 = axes[1, 0]
    x = tf["Location "]
    y = tf["Voltage-NE (V)"]
    colors = ["#5bc0de", "#5cb85c"]
    bars = ax3.bar(x, y, color=colors)
    ax3.set_xlabel("Location")
    ax3.set_ylabel("Voltage-NE (V)")
    ax3.set_title("Location VS Voltage-NE (V)")
    for bar, value in zip(bars, y):
        height = bar.get_height()
        ax3.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}%",
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Fourth subplot - Bar graph for loading
    ax4 = axes[1, 1]
    x = tf["Location "]
    tf["loading"] = tf["Average Phase Current (A)"] / tf["Rated Phase Current (A)"].round(2)
    y = tf["loading"]
    colors = ["#5bc0de", "#5cb85c"]
    bars = ax4.bar(x, y, color=colors)
    ax4.set_ylabel("loading")
    ax4.set_xlabel("Location")
    ax4.set_title("Location VS loading")
    for bar, value in zip(bars, y):
        height = bar.get_height()
        ax4.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}%",
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Adjust spacing between subplots
    plt.tight_layout()

    # Save the combined plot to a BytesIO object
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def resc_combined_graph(jf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    x = jf["Conductor Type"] + ", " + jf["Conductor Length (m)"].astype(str)
    y = jf["Corrected Continuity Resistance (Ω)"]
    colors = ["#d9534f", "#5bc0de", "#5cb85c", "#428bca"]
    sorted_indices = np.argsort(y)  # Sort the indices based on y values
    x_sorted = [x[i] for i in sorted_indices]
    y_sorted = [y[i] for i in sorted_indices]
    bars = plt.bar(x_sorted, y_sorted, color=colors)
    plt.xlabel("Conductor Type and Conductor Length (m) ")
    plt.ylabel("Corrected Continuity Resistance (Ω)")
    plt.title(
        "Conductor Type and Conductor Length (m) (V) VS Corrected Continuity Resistance (Ω) "
    )
    for bar, value in zip(bars, y_sorted):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value:.2f}",  # Show only the value of 'y' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    jf["Result"] = resc_rang(jf)
    jf_counts = jf["Result"].value_counts()
    labels = jf_counts.index.tolist()
    values = jf_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.axis("equal")
    plt.title("Test Results")
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def func_ops_combined_graph(of):
    plt.figure(figsize=(16, 8))

    # bar graph
    plt.subplot(121)
    tct = of["Functional Check"].value_counts()
    tct1 = of["Interlock check"].value_counts()

    X = tct.index
    Ygirls = tct.values
    Zboys = tct1.values

    X_axis = np.arange(len(X))

    plt.bar(X_axis - 0.2, Ygirls, 0.4, label="ok")
    plt.bar(X_axis + 0.2, Zboys, 0.4, label="notok")

    plt.xticks(X_axis, X)
    plt.xlabel("Groups")
    plt.ylabel("Number of Students")
    plt.title("Number of Students in each group")
    plt.legend()
    plt.show()

    # Pie chart
    plt.subplot(122)
    of["Result"] = func_ops_rang(of)
    of_counts = of["Result"].value_counts()
    labels = of_counts.index.tolist()
    values = of_counts.values.tolist()
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
    plt.axis("equal")
    plt.title("Test Results")
    # Save the combined graph as bytes
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def pat_combined_graph(bf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    combined_data = pd.concat(
        [bf["Functional Check"], bf["Visual Inspection"]], keys=["col1", "col2"]
    )
    tct = combined_data.value_counts()
    x_values = tct.index
    y_values = tct.values
    colors = ["#b967ff", "#e0a899"]
    bars = plt.bar(x_values, y_values, color=colors)
    plt.xlabel("functional and inspection check")
    plt.ylabel("Count")
    plt.title("Residual Test Results")

    # Add text labels inside each bar
    for bar, value in zip(bars, y_values):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,
            f"{value}",  # Show the count of 'col1' inside the bar
            ha="center",
            va="center",
            color="black",
            fontsize=8,
            rotation=0,
        )

    # Pie chart
    plt.subplot(122)
    result_counts = bf["Overall Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values
    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", shadow=False, startangle=90, colors=colors)
    plt.title("Residual Test Results")
    plt.axis("equal")  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def socket_combined_graph(sf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    sf_max_eli = sf.groupby("Location")[["L1-ELI", "L2-ELI", "L3-ELI"]].max()
    locations = sf_max_eli.index
    max_eli_values = sf_max_eli.max(axis=1)
    max_eli_locations = sf_max_eli.idxmax(
        axis=1
    )  # Get the locations corresponding to the max values
    colors = ["#d9534f", "#5bc0de", "#5cb85c", "#428bca"]
    bars = plt.bar(locations, max_eli_values, color=colors)
    plt.xlabel("Location")
    plt.ylabel("Max ELI (Ω)")
    plt.title("Max ELI for each Location")
    plt.xticks(rotation=90)

    # Add values inside each bar vertically
    for bar, location, value in zip(bars, max_eli_locations, max_eli_values):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,  # Adjust the y-coordinate to place text in the middle of the bar
            f"{value:.2f}\n({location})",
            ha="center",
            va="center",
            color="black",  # Set text color to white for better visibility on colored bars
            fontsize=8,  # Adjust font size as needed
            rotation=90,  # Set rotation to 90 degrees for vertical text
        )

    # Set the y-axis to a logarithmic scale
    plt.yscale("log")

    # Pie chart
    plt.subplot(122)
    result_counts = sf2["Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values

    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", shadow=False, startangle=90, colors=colors)
    plt.title("Residual Test Results")
    plt.axis("equal")  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined


def eli_test_combined_graph(gf):
    plt.figure(figsize=(16, 8))

    # Bar graph
    plt.subplot(121)
    gf_max_eli = gf.groupby("Location")[["L1-ELI", "L2-ELI", "L3-ELI"]].max()
    locations = gf_max_eli.index
    max_eli_values = gf_max_eli.max(axis=1)
    max_eli_locations = gf_max_eli.idxmax(
        axis=1
    )  # Get the locations corresponding to the max values
    colors = ["#d9534f", "#5bc0de", "#5cb85c", "#428bca"]
    bars = plt.bar(locations, max_eli_values, color=colors)
    plt.xlabel("Location")
    plt.ylabel("Max ELI (Ω)")
    plt.title("Max ELI for each Location")
    plt.xticks(rotation=90)

    # Add values inside each bar vertically
    for bar, location, value in zip(bars, max_eli_locations, max_eli_values):
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2,  # Adjust the y-coordinate to place text in the middle of the bar
            f"{value:.2f}\n({location})",
            ha="center",
            va="center",
            color="black",  # Set text color to white for better visibility on colored bars
            fontsize=8,  # Adjust font size as needed
            rotation=90,  # Set rotation to 90 degrees for vertical text
        )

    # Set the y-axis to a logarithmic scale
    plt.yscale("log")

    # Pie chart
    plt.subplot(122)
    result_counts = gf2["Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values

    colors = ["#00FF00", "#FF0000"]
    plt.pie(values, labels=labels, autopct="%1.1f%%", shadow=False, startangle=90, colors=colors)
    plt.title("Residual Test Results")
    plt.axis("equal")  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph_combined = io.BytesIO()
    plt.savefig(graph_combined)
    plt.close()

    return graph_combined
