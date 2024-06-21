
def format_solution(data, manager, routing, solution):

    formattedSolution = []

    time_dimension = routing.GetDimensionOrDie("Time")
    print("got to format solution")

    for vehicle_id in range(data["num_vehicles"]):

        formattedSolution.append({"pickupLocations": [], "dropoffLocations": [], "pickupTimes": [], "dropoffTimes": [], "vehicle" : data["vehicles"][vehicle_id], "passengers": [], "driver": data["drivers"][vehicle_id], "shipName": [], "tourName": [], "dockRep": []})

        index = routing.Start(vehicle_id)
        plan_output = f"Route of vehicle {data["vehicles"][vehicle_id]}:\n"
        iterations = 0
        while not routing.IsEnd(index):
            time_var = time_dimension.CumulVar(index)

            if iterations != 0:
                if iterations % 2 != 0:
                    formattedSolution[-1]["pickupLocations"].append(data["locations"][manager.IndexToNode(index)])
                    formattedSolution[-1]["pickupTimes"].append(data["original_times"][manager.IndexToNode(index)])
                    formattedSolution[-1]["passengers"].append(data["passenger_counts"][manager.IndexToNode(index)])
                    formattedSolution[-1]["shipName"].append(data["shipNames"][manager.IndexToNode(index)])
                    formattedSolution[-1]["tourName"].append(data["tourNames"][manager.IndexToNode(index)])
                    formattedSolution[-1]["dockRep"].append(data["dockReps"][manager.IndexToNode(index)])
                else:
                    formattedSolution[-1]["dropoffLocations"].append(data["locations"][manager.IndexToNode(index)])
                    formattedSolution[-1]["dropoffTimes"].append(data["original_times"][manager.IndexToNode(index)])
            iterations += 1
            

            plan_output += (
                f"{data["locations"][manager.IndexToNode(index)]} {data["original_times"][manager.IndexToNode(index)]}"
                f" Time({solution.Min(time_var)}, {solution.Max(time_var)})"
                " -> "
            )
            index = solution.Value(routing.NextVar(index))
        time_var = time_dimension.CumulVar(index)
        plan_output += (
            f"{data["locations"][manager.IndexToNode(index)]}"
            f" Time({solution.Min(time_var)}, {solution.Max(time_var)}min\n)"
        )
        plan_output += f"Time of the route: {solution.Min(time_var)}min\n"
        print(plan_output)
    
    return formattedSolution





