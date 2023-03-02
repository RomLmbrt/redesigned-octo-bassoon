#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

# new file created
touch("results.csv")

# file handling in write mode
f = open("results.csv", "w")

names = XLSX.readdata(data_file, "input_data", "J4:J8")
for name in names
    write(f, "$name ;")
end
write(f, "Hydro ; STEP pompage ; STEP turbinage ; Batterie injection ; Batterie soutirage ; Wind ;  RES ; load ; Net load ;  Cout Carbonne ; Nombre Parcs Installés \n")


Nweek = 1 #Number of week for the simulation

Tmax = 168 #optimization for 1 week (7*24=168 hours)
data_file = "data.xlsx"
for i in 1:Nweek


    td = 4 + (i-1)*168
    tf = 171 + (i-1)*168
    load = XLSX.readdata(data_file, "input_data", "C$(td):C$(tf)")
    wind = XLSX.readdata(data_file, "input_data", "D$(td):D$(tf)")
    solar = XLSX.readdata(data_file, "input_data", "E$(td):E$(tf)")
    hydro_fatal = XLSX.readdata(data_file, "input_data", "F$(td):F$(tf)")
    thermal_fatal = XLSX.readdata(data_file, "input_data", "G$(td):G$(tf)")
    #total of RES excepted wind
    Pres = solar + hydro_fatal + thermal_fatal #+ wind


#data for thermal clusters
Nth = 5 #number of thermal generation units
#names = XLSX.readdata(data_file, "input_data", "J4:J8")
dict_th = Dict(i=> names[i] for i in 1:Nth)
costs_th = XLSX.readdata(data_file, "input_data", "K4:K8")
Pmin_th = XLSX.readdata(data_file, "input_data", "M4:M8") #MW
Pmax_th = XLSX.readdata(data_file, "input_data", "L4:L8") #MW
dmin = XLSX.readdata(data_file, "input_data", "N4:N8") #hours

#data for hydro reservoir
Nhy = 1 #number of hydro generation units
Pmin_hy = zeros(Nhy)
Pmax_hy = XLSX.readdata(data_file, "input_data", "R4") *ones(Nhy) #MW
e_hy = XLSX.readdata(data_file, "input_data", "S4")*ones(Nhy) #MWh

#costs
cth = repeat(costs_th', Tmax) #cost of thermal generation €/MWh
chy = repeat([0], Tmax) #cost of hydro generation €/MWh
cuns = repeat([5000],Tmax) #cost of unsupplied energy €/MWh
cexc = repeat([5000], Tmax) #cost of in excess energy €/MWh


#data for STEP/battery
#weekly STEP
Pmax_STEP = 1200 #MW
rSTEP = 0.75

#battery
Pmax_battery = 280 #MW
rbattery = 0.85
d_battery = 2 #hours

#Consumption carbon print gCO2/kWh
cost_carbone_th = XLSX.readdata(data_file, "data", "C3:C7")
cost_carbone_hy = XLSX.readdata(data_file, "data", "C8:C8")
c_carbone_th = repeat(cost_carbone_th', Tmax)#cost of thermal generation gCO2/kWh
c_carbone_hy = repeat(cost_carbone_hy, Tmax)#cost of hydro generation gCO2/kWh

cost_carbone_wind = XLSX.readdata(data_file, "data", "C10:C10")
c_carbone_wind = repeat(cost_carbone_wind', Tmax)#cost of wind generation gCO2/kWh
cost_carbone_solar = XLSX.readdata(data_file, "data", "C11:C11")
c_carbone_solar = repeat(cost_carbone_solar', Tmax)#cost of solar generation gCO2/kWh

cost_carbone_th_fat = XLSX.readdata(data_file, "data", "C12:C12")
c_carbone_th_fat = repeat(cost_carbone_th_fat', Tmax)#cost of thermal fatal generation gCO2/kWh
cost_carbone_battery = XLSX.readdata(data_file, "data", "C13:C13")
c_carbone_battery =  repeat(cost_carbone_battery', Tmax)#cost of battery generation gCO2/kWh

#Normalize Carbon/cost
French_current_carbon_mix = 60 #gCO2/kHw - RTE data
Norm_carbon = maximum(load) * French_current_carbon_mix

French_current_cost_mix = 13.1 #euros/kHw - RTE data
Norm_cost = maximum(load) * French_current_cost_mix

#data for wind
Nwindinit = 2000/6 #Initial number of wind turbine

wind = XLSX.readdata(data_file, "input_data", "D4:D171")
Pprodparc = 15 #Mw

capexwind = repeat([1700000/(25*12*30*24)],Tmax) #capex eolien €/MWh. Average btw different type of wind production (offshore, onshore, ...)
opexwind = repeat([58000/(12*30*24)],Tmax) #opex horaire eolien €/Mw.

#############################
#create the optimization model
#############################
model = Model(HiGHS.Optimizer)

#############################
#define the variables
#############################
#thermal generation variables
@variable(model, Pth[1:Tmax,1:Nth] >= 0)
@variable(model, UCth[1:Tmax,1:Nth], Bin)
@variable(model, UPth[1:Tmax,1:Nth], Bin)
@variable(model, DOth[1:Tmax,1:Nth], Bin)
#hydro generation variables
@variable(model, Phy[1:Tmax,1:Nhy] >= 0)
#unsupplied energy variables
@variable(model, Puns[1:Tmax] >= 0)
#in excess energy variables
@variable(model, Pexc[1:Tmax] >= 0)
#weekly STEP variables
@variable(model, Pcharge_STEP[1:Tmax] >= 0)
@variable(model, Pdecharge_STEP[1:Tmax] >= 0)
@variable(model, stock_STEP[1:Tmax] >= 0)
#battery variables
@variable(model, Pcharge_battery[1:Tmax] >= 0)
@variable(model, Pdecharge_battery[1:Tmax] >= 0)
@variable(model, stock_battery[1:Tmax] >= 0)
#Wind variables
@variable(model, Pwind[1:Tmax] >= 0)
@variable(model, Nwind[1:Tmax] >= Nwindinit)


#cost variables
#Carbone_cost = sum((Pth.*c_carbone_th)[h] for h in 1:Nth)  + sum((Phy.*c_carbone_hy)[h] for h in 1:Nhy) + sum(wind.*c_carbone_wind) + sum(solar.*c_carbone_solar) + sum(hydro_fatal.*c_carbone_hy) + sum(thermal_fatal.*c_carbone_th_fat) + sum((Pcharge_battery+Pdecharge_battery).*c_carbone_battery)
#Financial_cost = sum((Pth.*cth)[h] for h in 1:Nth)+sum((Phy.*chy)[h] for h in 1:Nhy) + Puns'cuns + Pexc'cexc

#
# #############################
#define the objective function
#############################
#Last objective function, changed for a new one where the user chooses the number of wind turbines he/she wants
@objective(model, Min, (sum(Pth.*c_carbone_th)  + sum(Phy.*c_carbone_hy) + sum(Pwind.*c_carbone_wind) + sum(solar.*c_carbone_solar) + sum(hydro_fatal.*c_carbone_hy) + sum(thermal_fatal.*c_carbone_th_fat) + sum((Pcharge_battery+Pdecharge_battery).*c_carbone_battery))/Norm_carbon + (sum(Pth.*cth)+sum(Phy.*chy)+Puns'cuns+Pexc'cexc+sum((opexwind+capexwind)*(Pwind-wind)'))/Norm_cost )

#@objective(model, Min, (sum(Pth.*c_carbone_th)  + sum(Phy.*c_carbone_hy) + sum(Pwind.*c_carbone_wind) + sum(solar.*c_carbone_solar) + sum(hydro_fatal.*c_carbone_hy) + sum(thermal_fatal.*c_carbone_th_fat) + sum((Pcharge_battery+Pdecharge_battery).*c_carbone_battery))/Norm_carbon + (sum(Pth.*cth)+sum(Phy.*chy)+Puns'cuns+Pexc'cexc)/Norm_cost )


#############################
#define the constraints
#############################
#balance constraint
#@constraint(model, balance[t in 1:Tmax], sum(Pth[t,g] for g in 1:Nth) + sum(Phy[t,h] for h in 1:Nhy) + Pres[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)
@constraint(model, balance[t in 1:Tmax], sum(Pth[t,g] for g in 1:Nth) + sum(Phy[t,h] for h in 1:Nhy) + Pwind[t] + Pres[t] - load[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)

#thermal unit Pmax constraints
@constraint(model, max_th[t in 1:Tmax, g in 1:Nth], Pth[t,g] <= Pmax_th[g]*UCth[t,g])
#thermal unit Pmin constraints
@constraint(model, min_th[t in 1:Tmax, g in 1:Nth], Pmin_th[g]*UCth[t,g] <= Pth[t,g])
#thermal unit Dmin constraints
for g in 1:Nth
        if (dmin[g] > 1)
            @constraint(model, [t in 2:Tmax], UCth[t,g]-UCth[t-1,g]==UPth[t,g]-DOth[t,g],  base_name = "fct_th_$g")
            @constraint(model, [t in 1:Tmax], UPth[t]+DOth[t]<=1,  base_name = "UPDOth_$g")
            @constraint(model, UPth[1,g]==0,  base_name = "iniUPth_$g")
            @constraint(model, DOth[1,g]==0,  base_name = "iniDOth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] >= sum(UPth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminUPth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] <= 1 - sum(DOth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminDOth_$g")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] >= sum(UPth[i,g] for i in 1:t), base_name = "dminUPth_$(g)_init")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] <= 1-sum(DOth[i,g] for i in 1:t), base_name = "dminDOth_$(g)_init")
    end
end

#hydro unit constraints
@constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t,h] <= Pmax_hy[h])

#hydro stock constraint
@constraint(model, capacity_hy[h in 1:Nhy], sum(Phy[t,h] for t in 1:Tmax) <= e_hy[h])

#weekly STEP
@constraint(model, max_charge_step[t in 1:Tmax], Pcharge_STEP[t] <= 1200)
@constraint(model, max_decharge_step[t in 1:Tmax], Pdecharge_STEP[t] <= 1200)
@constraint(model, limit_stock_step[t in 1:Tmax], stock_STEP[t] <=  1200*Tmax)
#hypothèse : On part d'un stock initiale à 50% d'une capaité totale supposée égale à (Nombre heures semaine)*Puissance
@constraint(model, init_stock_step, stock_STEP[1] == 1200*Tmax*0.5)
@constraint(model, periode_stock_step, stock_STEP[1] == stock_STEP[Tmax])
@constraint(model, stock_limit_step[t in 2:Tmax], stock_STEP[t]==stock_STEP[t-1]+Pcharge_STEP[t-1]*0.75-Pdecharge_STEP[t-1])

#battery
@constraint(model, max_charge_battery[t in 1:Tmax], Pcharge_battery[t] <= 280)
@constraint(model, max_decharge_battery[t in 1:Tmax], Pdecharge_battery[t] <= 280)
@constraint(model, limit_stock_battery[t in 1:Tmax], stock_battery[t] <=  280*2)

@constraint(model, init_stock_battery, stock_battery[1] == 0)
#hypothèse : on considère les puissances définis par rapport au réseau
@constraint(model, stock_limit_battery[t in 2:Tmax], stock_battery[t]==stock_battery[t-1]+Pcharge_battery[t]*0.85-Pdecharge_battery[t]/0.85)

#Wind
@constraint(model, max_puissance_wind[t in 1:Tmax], Pwind[t] <= (Nwind[t]/Nwindinit + 1 ) * wind[t])
@constraint(model, min_puissance_wind[t in 1:Tmax], Pwind[t] >= (Nwind[t]/Nwindinit + 1 ) * wind[t])
@constraint(model, nwind_increase[t in 2:Tmax], Nwind[t]>=Nwind[t-1])


#no need to print the model when it is too big
#solve the model
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)
#@show Financial_cost


#exports results as csv file
th_gen = value.(Pth)
hy_gen = value.(Phy)
wind_gen = value.(Pwind)
Nwind_gen = value.(Nwind)
STEP_charge = value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)
battery_charge = value.(Pcharge_battery)
battery_decharge = value.(Pdecharge_battery)

for t in 1:Tmax
    for g in 1:Nth
        write(f, "$(th_gen[t,g]) ; ")
    end
    for h in 1:Nhy
        write(f, "$(hy_gen[t,h]) ;")
    end
    write(f, "$(STEP_charge[t]) ; $(STEP_decharge[t]) ;")
    write(f, "$(battery_charge[t]) ; $(battery_decharge[t]) ;")
    cout_th = zeros(Tmax,1)
    cout_hy = zeros(Tmax,1)
    for h in 1:Nth
        cout_th[t] += th_gen[t,h]*c_carbone_th[h]   
    end
    for h in 1:Nhy
        cout_hy[t] += hy_gen[t,h]*c_carbone_hy[h]
    end
    carbone_cost = cout_th[t] + cout_hy[t] + (c_carbone_wind[t]*wind[t]) + (solar[t]*c_carbone_solar[t]) + (hydro_fatal[t]*c_carbone_hy[t]) + (thermal_fatal[t]*c_carbone_th_fat[t]) + (battery_charge[t]+battery_decharge[t])*c_carbone_battery[t]
    #Financial_cost = sum((Pth.*cth)[h] for h in 1:Nth)[t]+sum((Phy.*chy)[h] for h in 1:Nhy) + Puns'cuns + Pexc'cexc
    write(f, "$(wind_gen[t]) ; $(Pres[t]) ;  $(load[t]) ; $(load[t]-Pres[t]) ; $(carbone_cost[1]) ; $(Nwind_gen[t] - Nwindinit) \n")

end
end

close(f)
