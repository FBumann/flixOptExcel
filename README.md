# README
## Introduction
**flixOptExcel** is an Extension to the optimization framework [flixOpt](https://github.com/flixOpt/flixOpt), tailored for district heating systems. 
It uses excel for setting parameters and therefore makes optimization accessible to a broad range of users.
The automated and comprehensive evaluation process enables quick understanding of the computed results.

**Table of Contents**
- [Introduction](#introduction)
- [Code Structure](#general-code-structure)
- [Setup and update](#setup-the-project)
- [District Heating](#district-heating)
- [Modeling of the District Heating System](#modeling-of-the-district-heating-system)
	- [Heat Demand and Heating Network](#heat-demand-and-heating-network)
		- [Computation of Heating Losses of the District Heating Network](#computation-of-heating-losses-of-the-district-heating-network)
	- [Components](#components)
		- [General](#general)
		- [Investments](#investments)
		- [Technologies](#technologies)
			- [CHP (Combined Heat and Power)](#chp-combined-heat-and-power)
			- [Boiler](#boiler)
			- [Electrode Boiler](#electrode-boiler)
			- [Waste Heat](#waste-heat)
			- [Heat Pump](#heat-pump)
			- [Waste Heat Pump](#waste-heat-pump)
			- [Storage](#storage)
			- [Cooling Tower](#cooling-tower)
			- [CHP with variable rate between electricity and heat](#chp-with-variable-rate-between-electricity-and-heat)
	- [Additional Information](#additional-information)
		- [Energy Inputs](#energy-inputs)
			- [Fuel](#fuel)
			- [Others](#others)
		- [Energy Outputs](#energy-outputs)
		- [Operation Funding](#operation-funding)
		- [Values Changing over Time](#values-changing-over-time)

## Setup the Project
1. Create a new Python project in your IDE (PyCharm, Spyder, ...) and activate it
2. Install this package via pip in to your environment: `pip install git+https://github.com/FBumann/flixOptExcel.git`
3. Copy the "main.py" file from this package into your project.
4. Make a local copy of the Template_DataInput.xlsx and save it somewhere on your Computer (for. ex. Desktop)
5. Copy the path of the excel file into the "main.py" file
6. **Edit the Excel-file to initialize your Model.**
7. Run the main.py file
8. Analyse the results of your Model. It's saved under the path you specified in the input-excel-file

To update the repository to the newest version, simply do a clean uninstall and reinstall:
`pip uninstall flixOptExcel `
`pip install --upgrade git+https://github.com/FBumann/flixOptExcel.git@main`

## District Heating
A District heating system is used to supply customers with heat for housing and hot water, but also industries with heat for industrial processes. This heat is mostly produced in big plants and transported in district heating networks - mostly pipes containing hot water.
**flixOptExcel** provides an easy to use interface to create a digital copy of such a system and optimize its operation. Furthermore, it can be used to evaluate different Investment options and support evidence based decisions.
## Modeling of the District Heating System
The District heating system is modeled with one-hour time-increments. It's total Heat Demand per time step has to be specified. All available and possible Components to produce this heat also need to be specified by the user. The Model optimizes the Operation - and, if provided, Investments - of these Components, to produce the specified heat demand with minimal costs.
### Heat Demand and Heating Network
The Heat Demand of a district heating system consists of the heat demand of all consumers and the losses of the heating network. The combined heat demand of all consumers per time step needs to be specified in the **Zeitreihen Sheets**. Internal [Computation of Heating Losses of the District Heating Network](#computation-of-heating-losses-of-the-district-heating-network) is possible.
#### Computation of Heating Losses of the District Heating Network
Description of the logic
### Components
#### General
Components are modeled as Linear Transformers, having **[Energy Inputs](#Energy-Inputs)** and **[Energy Outputs](#Energy-Outputs)** with a fixed relation between them, called **key performance indicator** (KPI). Further, the thermal performance of every component has to be defined.
KPI's and other values can be [Values Changing over Time](#values-changing-over-time).
- `Name: 'unique_name'` The name of the Component. Has to be unique.
Optional:
- `group: 'group_name'` Assign the Component to a group. Only affects Evaluation
- `Thermische Leistung: 20` Thermal Power in MW<sub>th</sub>. To Optimize, either leave blank or give boundaries (`Thermische Leistung: '20-40'`). See [Investments](#Investments).
#### Investments
Optimizing Investments is a key part of [flixOpt](https://github.com/flixOpt/flixOpt). Therefore, all components have the option to optimize their thermal power.
To optimize thermal power, either leave `Thermische Leistung` blank or give boundaries (`Thermische Leistung: '20-40'`)
**Arguments:**
- `Optional: ja/nein` Defines, wether the investment is forced or not
- `Fixkosten pro MW und Jahr: 40` Fixed costs per year in €/MW<sub>th</sub> 
Optional:
- `Fixkosten pro Jahr: 30` Fixed costs per year in €
- `Förderung pro Jahr: 10` Fixed costs per year in €
- `Förderung pro MW und Jahr: 20` Funding per year in €/MW<sub>th</sub>
- `Startjahr: 2020` First year of operation of the component
- `Endjahr: 2040` Last year of Operation of the component (included)
- `Invest Gruppe: 'groupname:40'` Used to limit investments for multiple Components at once. The total thermal power of all Components having this attribute can not exceed 40.
#### Technologies
Following Technologies are currently supported. Suggestions are happily welcomed.
- [CHP (Combined Heat and Power)](#chp-combined-heat-and-power)
- [Boiler](#boiler)
- [Electrode Boiler](#electrode-boiler)
- [Waste Heat](#waste-heat)
- [Heat Pump](#heat-pump)
- [Waste Heat Pump](#waste-heat-pump)
- [Storage](#storage)
- [Cooling Tower](#cooling-tower)
- [CHP with variable rate between electricity and heat](#chp-with-variable-rate-between-electricity-and-heat)
##### CHP (Combined Heat and Power)
Uses [Fuel](#Fuel) to produce `Heat`and `Electricity`. Has 2 KPI's: thermal efficiency (eta<sub>th</sub>) and electric efficiency (eta<sub>el</sub>). 
- `eta_th: 40%` thermal efficiency
- `eta_el: 40%` electric efficiency
- `Brennstoff: 'Erdgas'`Type of [Fuel](#Fuel) used
Optional:
- `Zusatzkosten pro MWh Brennstoff: 4` Extra costs in €/MWh<sub>fuel</sub> on top of regular costs for [Fuel](#Fuel)
##### Boiler
Uses [Fuel](#Fuel) to produce `Heat`. Has 1 KPI: thermal efficiency (eta<sub>th</sub>)
- `eta_th: 85%` thermal efficiency
- `Brennstoff: 'Erdgas'` Type of [Fuel](#Fuel) used
*Optional:*
- `Zusatzkosten pro MWh Brennstoff: 4` Extra costs in €/MWh<sub>fuel</sub> on top of regular costs for [Fuel](#Fuel)
##### Electrode Boiler
Uses `Electricity` to produce `Heat`. Has 1 KPI: thermal efficiency (eta<sub>th</sub>)
- `eta_th: 85%` thermal efficiency
*Optional:*
- `Zusatzkosten pro MWh Strom: 4` Extra costs in €/MWh<sub>el</sub> on top of regular costs for [Electricity](#Energy-Inputs)
##### Waste Heat 
Uses `Waste Heat`to produce `Heat`. Has no KPI.
Optional:
-  `Abwärmekosten: 20` Costs for excess heat in €/MWh<sub>amb</sub>
##### Heat Pump
Uses `Electricity`to produce `Heat`. Has 1 KPI: thermal efficiency (COP)
- `COP: 3` thermal efficiency
- `COP berechnen: ja` wether to calculate the COP internally. specify the source temperature in `COP`
Optional:
- `Zusatzkosten pro MWh Strom: 4` Extra costs in €/MWh<sub>el</sub> on top of regular costs for [Electricity](#Energy-Inputs)
- `Untergrenze für Einsatz: 5` Limit for usage. Usage is prohibited for time steps, in which the time series specified in `Zeitreihe für Einsatzbeschränkung`is below this value
- `Zeitreihe für Einsatzbeschränkung: 'name_of_ts'` Name of time series for limiting useage
- `Betriebskostenförderung BEW: 92` Operation Funding as stated in the German Federal Program BEW in €/MWh<sub>ambient_heat</sub>.
##### Waste Heat Pump
Uses `Electricity` and `Waste Heat`to produce `Heat`. Has 1 KPI: thermal efficiency (COP)
- `COP: 3` thermal efficiency
- `COP berechnen: ja` wether to calculate the COP internally. specify the source temperature in `COP`
Optional:
-  `Abwärmekosten: 4` Costs for excess heat in €/MWh<sub>amb</sub>
- `Zusatzkosten pro MWh Strom: 4` Extra costs in €/MWh<sub>el</sub> on top of regular costs for [Electricity](#Energy-Inputs)
- `Untergrenze für Einsatz: 5` Limit for useage. Useage is prohibited for timesteps, in which the time series specified in `Zeitreihe für Einsatzbeschränkung`is below this value
- `Zeitreihe für Einsatzbeschränkung`
- `Betriebskostenförderung BEW: 92` [Operation Funding](#Operation-Funding) in €/MWh<sub>ambient_heat</sub>
##### Storage
Uses `Heat`, stores it and produces heat again. Has 3 KPI: Thermal efficiency for charging (eta_load), Thermal efficiency for discharging (eta_unload) and Heat Losses per hour (VerlustProStunde). Furthermore, has a thermal capacity, which can be optimized as well.
- `Kapazität [MWh]: 1000` Capacity in MWh<sub>th</sub> of the Storage
Optional:
- `VerlustProStunde: 0,5%` Heat Losses per hour relative to the current charge state in %/hour
- `eta_load: 98%` Thermal Efficiency of charging the storage
- `eta_unload: 98%` Thermal Efficiency of discharging the storage
- `AbhängigkeitVonDT: nein` Wether the Capacity and thermal Power of the storage should be limited according to the temperature spread in the district heating network. Maximum capacity at $dT = 65 K$.
**Additional For Investment of Capacity:**
- `Fixkosten pro MWh und Jahr: 10` Fixed costs per year in €/MWh<sub>th</sub>
- `Förderung pro MWh und Jahr: 20` Funding per year in €/MWh<sub>th</sub>

##### Cooling Tower

##### CHP with variable rate between electricity and heat

### Additional Information

#### Energy Inputs
Energy Inputs are endless Energy sources, only limited by price (€/MWh). Prices can vary between time steps. The prices need to be specified in the **Zeitreihen Sheets**.
##### Fuel
- `Erdgas`
- `Wasserstoff`
- `EBS`
##### Others
- `Waste Heat`
- `Electricity`
- `Heat`
#### Energy Outputs
Energy Outputs are either endless sinks with a price reward (€/MWh), or a Energy Demand, like Heat. The prices need to be specified in the **Zeitreihen Sheets**. Use negative values for rewards
- `Heat`A fixed amount of heat needs to be produced per time step. See [Heat Demand and Heating Network](#Heat-Demand-and-Heating-Network)
- `Electricity`
#### Operation Funding
The german federal Program BEW funds Heat Pumps with a fixed amount per MWh<sub>ambient</sub>, 
calculated by a complex formula which includes the SCOP. The funding is limited to a max of 90% of electricity costs, 
to the first 10 years of operation and further by the "Wirtschaftlichkeitslücke".

**Modeling**
- Funding in €/MWh<sub>ambient</sub> is estimated by user
- Assumption: Heat Pump reaches the SCOP of 2.5
- Funding is limited to first 10 years of operation
- Funding is applied per timestep and limited to 90% of electricity costs per time step.

This Approach applies the funding in a way, that the funding phase out after 10 years realistcally impacts operation, 
while giving the user maximum control of the amount of funding he wants to apply.

#### Values Changing over Time
KPI's and Prices can change over time. To use this feature, instead of assigning a number, pass a column name and create this column in the corresponding excel sheets.