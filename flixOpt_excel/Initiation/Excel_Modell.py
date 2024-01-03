# -*- coding: utf-8 -*-
import shutil
import random
import os
import openpyxl
import datetime

import pandas as pd
from pprintpp import pprint as pp
from typing import Literal, List

from flixOpt_excel.Initiation.HelperFcts_in import *
from flixOpt.flixComps import *
from flixOpt_excel.Evaluation import graphics_excel


def run_excel_model(excel_file_path: str, solver_name: str, gap_frac: float = 0.001, timelimit: int = 3600):
    """

    :param excel_file_path:
        path to the excel_file where the input data is stored
    :param solver_name:
        choose from "cbc", "gurobi", "glpk"
    :param gap_frac:
        0...1 ; gap to relaxed solution, kind of like "accuracy. Higher values for faster solving. 0...1
    :param timelimit:
        timelimit in seconds. After this time limit is exceeded the solution process is stopped
    """

    # Pfad der Input-Excel
    excel_file_path = excel_file_path

    # <editor-fold desc="Konstanten">
    t_co2_per_MWh_gas = 0.202
    # </editor-fold>
    # <editor-fold desc="Solver Inputs">
    solver_props = {'gapFrac': gap_frac,  # solver-gap
                   'timelimit': timelimit,  # seconds until solver abort
                   'solver': solver_name,
                   'displaySolverOutput': True,  # ausführlicher Solver-resources.
                   }
    # </editor-fold>

    # <editor-fold desc="Excel Import und checks">
    # <editor-fold desc="Allgemeines">
    allgemeine_info = pd.read_excel(excel_file_path, sheet_name="Allgemeines")
    years = [year for year in allgemeine_info["Jahre"] if isinstance(year, int)]
    results_directory = allgemeine_info.at[0, "Speicherort"] if isinstance(allgemeine_info.at[0, "Speicherort"],
                                                                          str) else None
    calc_name = allgemeine_info.at[0, "Name"] if isinstance(allgemeine_info.at[0, "Name"], str) else None

    co2_limit = allgemeine_info["CO2-limit"].replace({np.nan: None}, )[:len(years)]
    co2_limit_dict = dict(zip(years, co2_limit))

    use_portfolio = allgemeine_info.at[0, "Bestandsanlagen nutzen"] if allgemeine_info.at[
                                                                           0, "Bestandsanlagen nutzen"] in ["ja",
                                                                                                            "nein"] else None
    if any((results_directory, calc_name, use_portfolio)) is None or len(years) < 1:
        raise UserWarning("Bei Input 'allgemeines' müssen alle Felder ausgefüllt sein.")  # TODO: Fehermeldung verbessern

    use_portfolio = [{'ja': True, 'nein': False}[k] for k in [use_portfolio]][0]

    # </editor-fold>
    # <editor-fold desc="Zeitreihen">
    li = []  # Initialize an empty list to store DataFrames
    for i in range(len(years)):  # Iterate over the range of years
        sheet_name = f"Preiszeitreihen ({i + 1})" if i > 0 else "Preiszeitreihen"
        # Read the Excel sheet, skipping the first two rows, and drop specified columns
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=[1, 2]).drop(columns=["Tag", "Uhrzeit"])
        li.append(df)  # Append the DataFrame to the list
    preiszeitreihen = pd.concat(li, axis=0, ignore_index=True)  # Concatenate the DataFrames in the list


    li = []  # Initialize an empty list to store DataFrames
    for i in range(len(years)):  # Iterate over the range of years
        sheet_name = f"Sonstige Zeitreihen ({i + 1})" if i > 0 else "Sonstige Zeitreihen"
        # Read the Excel sheet, skipping the first two rows, and drop specified columns
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=[1, 2]).drop(columns=["Tag", "Uhrzeit"])
        li.append(df)  # Append the DataFrame to the list
    zeitreihen = pd.concat(li, axis=0, ignore_index=True)  # Concatenate the DataFrames in the list

    check_dataframe_consistency(df=preiszeitreihen, years=years, name_of_df="Preiszeitreihen")
    check_dataframe_consistency(df=zeitreihen, years=years, name_of_df="Zeitreihen")


    # <editor-fold desc="Erzeugerdaten">
    df = pd.read_excel(excel_file_path, sheet_name="Erzeuger", header=None)
    erzeugerdaten = handle_component_data(df)

    if use_portfolio:
        print("Bestandsanlagen werden genutzt.")
        df_portfolio = pd.read_excel(excel_file_path, sheet_name="Erzeuger_Bestand", header=None)
        erzeugerdaten_bestand = handle_component_data(df_portfolio)
        erzeugerdaten = combine_dicts_of_component_data(erzeugerdaten, erzeugerdaten_bestand)

    erzeugerdaten = convert_component_data_types(erzeugerdaten)
    erz_daten = convert_component_data_for_looping_through(erzeugerdaten)  # Keyword Zuweisung und Sortierung
    print("Erzeugerdaten wurden eingelesen.")
    # </editor-fold>
    # </editor-fold>

    # <editor-fold desc="Time Series creation and StartUp of EnergySystem">
    a_time_series = datetime.datetime(2021, 1, 1) + np.arange(len(zeitreihen.index)) * datetime.timedelta(hours=1)
    a_time_series = a_time_series.astype('datetime64')
    energy_system = cEnergySystem(a_time_series, dt_last=None) # creating System
    zeitreihen.index = a_time_series
    preiszeitreihen.index = a_time_series



    # <editor-fold desc="Fernwaermenetz (Verluste und Temperaturen)">
    zeitreihen['Tamb24mean'] = calculate_hourly_rolling_mean(series=zeitreihen['Tamb'], window_size=24)
    zeitreihen = handle_heating_network(zeitreihen)  # calculation f the heating network temperature and losses
    # </editor-fold>

    # <editor-fold desc="Busse">
    excess_costs = None
    b_strom_einspeisung = cBus('el', 'StromEinsp', excessCostsPerFlowHour=excess_costs)
    b_strom_bezug = cBus('el', 'StromBez', excessCostsPerFlowHour=excess_costs)
    b_fernwaerme = cBus('heat', 'Fernwaerme', excessCostsPerFlowHour=excess_costs)
    b_gas = cBus('fuel', 'Gas', excessCostsPerFlowHour=excess_costs)
    b_wasserstoff = cBus('fuel', 'Wasserstoff', excessCostsPerFlowHour=excess_costs)
    b_ebs = cBus(media='fuel', label='EBS', excessCostsPerFlowHour=excess_costs)
    b_abwaerme = cBus(media='heat', label='AbwaermeBus', excessCostsPerFlowHour=excess_costs)
    b_abwaerme_ht = cBus(media='heat', label='AbwaermeHTBus', excessCostsPerFlowHour=excess_costs)
    b_abwaerme_nt = cBus(media='heat', label='AbwaermeNTBus', excessCostsPerFlowHour=excess_costs)
    b_speicher_sum = cBus(media='heat', label='SpeicherBus', excessCostsPerFlowHour=excess_costs)

    # </editor-fold>
    # <editor-fold desc="Effekte">
    e_target = cEffectType('target', 'i.E.', 'Target',  # name, unit, description
                         isObjective=True)  # defining costs as objective of optimiziation
    e_costs = cEffectType('costs', '€', 'Kosten', isStandard=True,
                        specificShareToOtherEffects_operation={e_target: 1},
                        specificShareToOtherEffects_invest={e_target: 1})

    e_funding = cEffectType('funding', '€', 'Funding Gesamt',
                          specificShareToOtherEffects_operation={e_costs: -1},
                          specificShareToOtherEffects_invest={e_costs: -1})

    # Limit CO2 Emissions per year
    co2_limit_effects = []
    co2_limiter_shares = {}
    for year in years:
        if co2_limit_dict.get(year) is not None:
            co2_limit_yearly = cEffectType(f"CO2_limit_{year}", 't',
                                           description="Effect to limit the Emissions in that year",
                                           max_operationSum=co2_limit_dict[year])
            co2_limit_effects.append(co2_limit_yearly)
            exists = handle_operation_years({"Startjahr": year, "Endjahr": year}, years)
            co2_limiter_shares[co2_limit_yearly] = pd.DataFrame(exists)

    energy_system.addEffects(*co2_limit_effects)

    effect_shares_co2 = {e_costs: preiszeitreihen["costsCO2"]}
    effect_shares_co2.update(co2_limiter_shares)
    e_co2 = cEffectType('CO2', 't', 'CO2Emissionen',
                      specificShareToOtherEffects_operation=effect_shares_co2)

    energy_system.addEffects(e_costs, e_funding, e_co2, e_target)

    # </editor-fold>
    # <editor-fold desc="Sinks and Sources">
    s_waermelast = cSink('Waermelast', group="Wärmelast_mit_Verlust",
                       sink=cFlow('Qth', bus=b_fernwaerme, nominal_val=1, val_rel=zeitreihen["SinkHeat"]))
    energy_system.addComponents(s_waermelast)

    # Berechnungslogik und/oder InputInExcel
    s_netzverluste = cSink('Waermelast_Netzverluste', group="Wärmelast_mit_Verlust",
                         sink=cFlow('Qth', bus=b_fernwaerme, nominal_val=1, val_rel=zeitreihen["SinkLossHeat"]))
    energy_system.addComponents(s_netzverluste)

    s_strom_einspeisung = cSink('StromEinspeisung', sink=cFlow('Pel', bus=b_strom_einspeisung))
    energy_system.addComponents(s_strom_einspeisung)

    s_strom_bezug = cSink('StromBezug', sink=cFlow('Pel', bus=b_strom_bezug))
    energy_system.addComponents(s_strom_bezug)

    s_gas_bezug = cSource('GasBezug', source=cFlow('Qfu', bus=b_gas))
    energy_system.addComponents(s_gas_bezug)

    s_wasserstoff_bezug = cSource('WasserstoffBezug', source=cFlow('Qfu', bus=b_wasserstoff))
    energy_system.addComponents(s_wasserstoff_bezug)

    s_ebs_bezug = cSource('EBSBezug', source=cFlow('Qfu', bus=b_ebs))
    energy_system.addComponents(s_ebs_bezug)

    s_abwaerme_bezug = cSource(label="AbwaermeBezug", source=cFlow(label='Qabw', bus=b_abwaerme))
    energy_system.addComponents(s_abwaerme_bezug)

    # Hilfsbusse HT und NT
    Qin = cFlow(label="Qin", bus=b_abwaerme)
    Qout = cFlow(label="Qout", bus=b_abwaerme_ht)
    CompHT = cBaseLinearTransformer(label="HelperAbwaermeHT", inputs=[Qin], outputs=[Qout],
                                    factor_Sets=[{Qin: 1, Qout: 1}])

    Qin = cFlow(label="Qin", bus=b_abwaerme)
    Qout = cFlow(label="Qout", bus=b_abwaerme_nt)
    CompNT = cBaseLinearTransformer(label="HelperAbwaermeNT", inputs=[Qin], outputs=[Qout],
                                    factor_Sets=[{Qin: 1, Qout: 1}])
    energy_system.addComponents(CompNT, CompHT)
    # </editor-fold>

    # <editor-fold desc="KWK">
    if "KWK" in erz_daten:
        KWKs = list()
        for item in erz_daten["KWK"]:
            # Zuweisung Brennstoff und Netzentgelte
            # TODO: handle CO2 split for electricity and heat
            fuel_bus, fuel_costs = handle_fuel_input_switch_excel(item, preiszeitreihen, b_gas, b_wasserstoff, b_ebs, e_costs,
                                                                  e_co2, t_co2_per_MWh_gas)
            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aKWK = cKWK(label="KWK" + item["label"], exists=exists, group=item.get("group", None),
                        eta_th=item["eta_th"], eta_el=item["eta_el"],
                        Q_th=cFlow(label='Qth', bus=b_fernwaerme, nominal_val=nominal_val, investArgs=invest),
                        P_el=cFlow(label='Pel', bus=b_strom_einspeisung,
                                   costsPerFlowHour={e_costs: preiszeitreihen["costsEinspEl"]}),
                        Q_fu=cFlow(label='Qfu', bus=fuel_bus, costsPerFlowHour=fuel_costs)
                        )
            KWKs.append(aKWK)

        energy_system.addComponents(*KWKs)
    # </editor-fold>
    # <editor-fold desc="Kessel">
    if "Kessel" in erz_daten:
        Kessels = list()
        for item in erz_daten["Kessel"]:
            # Zuweisung Brennstoff und Netzentgelte
            fuel_bus, fuel_costs = handle_fuel_input_switch_excel(item, preiszeitreihen, b_gas, b_wasserstoff, b_ebs, e_costs,
                                                                  e_co2, t_co2_per_MWh_gas)
            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aKessel = cKessel(label="Kessel" + item["label"], exists=exists, group=item.get("group", None),
                              eta=item["eta_th"],
                              Q_th=cFlow(label='Qth', bus=b_fernwaerme, nominal_val=nominal_val, investArgs=invest),
                              Q_fu=cFlow(label='Qfu', bus=fuel_bus, costsPerFlowHour=fuel_costs)
                              )
            Kessels.append(aKessel)

        energy_system.addComponents(*Kessels)
    # </editor-fold>
    # <editor-fold desc="EHK">
    if "EHK" in erz_daten:
        EHKs = list()
        for item in erz_daten["EHK"]:
            BusInput = b_strom_bezug
            fuel_costs = {e_costs: preiszeitreihen["costsBezugEl"] + item.get("ZusatzkostenEnergieInput", 0)}

            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aEHK = cEHK(label="EHK" + item["label"], eta=item["eta_th"],
                        exists=exists, group=item.get("group", None),
                        Q_th=cFlow(label='Qth', bus=b_fernwaerme, exists=exists,
                                   nominal_val=nominal_val, investArgs=invest),
                        P_el=cFlow(label='Pel', bus=BusInput, costsPerFlowHour=fuel_costs)
                        )
            EHKs.append(aEHK)
        energy_system.addComponents(*EHKs)
    # </editor-fold>
    # <editor-fold desc="Waermepumpe">
    if "Waermepumpe" in erz_daten:
        WPs = list()
        for item in erz_daten["Waermepumpe"]:
            FuelcostsEl = {e_costs: preiszeitreihen["costsBezugEl"] + item.get("ZusatzkostenEnergieInput")}

            # COP-Berechnung
            COP = handle_COP_calculation(item, zeitreihen)

            # Betriebskostenförderung
            fund_op = handle_operation_fund_of_heatpump(item, e_funding, COP, FuelcostsEl[e_costs], 92)

            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aWP = cHeatPump(label="Waermepumpe" + item["label"], COP=COP,
                            exists=exists, group=item.get("group", None),
                            Q_th=cFlow(label='Qth', bus=b_fernwaerme,
                                       nominal_val=nominal_val, exists=exists,
                                       costsPerFlowHour=fund_op,
                                       investArgs=invest),
                            P_el=cFlow(label='Pel', bus=b_strom_bezug, costsPerFlowHour=FuelcostsEl)
                            )
            WPs.append(aWP)

        energy_system.addComponents(*WPs)
    # </editor-fold>
    # <editor-fold desc="TAB">
    if "TAB" in erz_daten:
        TABs = list()
        for item in erz_daten["TAB"]:
            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aTAB = cKWK(label="TAB" + item["label"], eta_th=item["eta_th"], eta_el=item["eta_el"],
                        exists=exists, group=item.get("group", None),
                        Q_th=cFlow(label='Qth', bus=b_fernwaerme, exists=exists,
                                   nominal_val=nominal_val, investArgs=invest),
                        P_el=cFlow(label='Pel', bus=b_strom_einspeisung,
                                   costsPerFlowHour={e_costs: preiszeitreihen["costsEinspEl"]}),
                        Q_fu=cFlow(label='Qfu', bus=b_ebs, costsPerFlowHour={e_costs: preiszeitreihen["costsBezugEBS"]})
                        )
            TABs.append(aTAB)
        energy_system.addComponents(*TABs)
    # </editor-fold>
    # <editor-fold desc="Abwaerme HT">
    if "AbwaermeHT" in erz_daten:
        AbwaermeHTs = list()
        for item in erz_daten["AbwaermeHT"]:
            BusInput = b_abwaerme_ht
            Fuelcosts = {e_costs: item["costsPerFlowHour_abw"]}

            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            Qin = cFlow(label="Qab", bus=b_abwaerme_ht, costsPerFlowHour=Fuelcosts)
            Qout = cFlow(label="Qth", bus=b_fernwaerme,
                         nominal_val=nominal_val, exists=exists,
                         investArgs=invest)

            aAbwaermeHT = cBaseLinearTransformer(label="AbwaermeHT" + item["label"],
                                                 exists=exists, group=item.get("group", None),
                                                 inputs=[Qin], outputs=[Qout], factor_Sets=[{Qin: 1, Qout: 1}])
            AbwaermeHTs.append(aAbwaermeHT)
        energy_system.addComponents(*AbwaermeHTs)
    # </editor-fold>
    # <editor-fold desc="Abwaerme WP">
    if "AbwaermeWP" in erz_daten:
        AbwaermeWPs = list()
        for item in erz_daten["AbwaermeWP"]:
            FuelcostsEl = {e_costs: preiszeitreihen["costsBezugEl"] + item.get("ZusatzkostenEnergieInput", 0)}
            FuelcostsAbw = {e_costs: item["costsPerFlowHour_abw"]}

            COP = handle_COP_calculation(item, zeitreihen)

            # Betriebskostenförderung
            fund_op = handle_operation_fund_of_heatpump(item, e_funding, COP, FuelcostsEl[e_costs], 92)

            # Investitionskosten und Förderung
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aAbwaermeWP = cAbwaermeHP(label="AbwaermeWP" + item["label"], COP=COP,
                                      exists=exists, group=item.get("group", None),
                                      Q_th=cFlow(label='Qth', bus=b_fernwaerme,
                                                 nominal_val=nominal_val, exists=exists,
                                                 investArgs=invest,
                                                 costsPerFlowHour=fund_op),
                                      P_el=cFlow(label='Pel', bus=b_strom_bezug, costsPerFlowHour=FuelcostsEl),
                                      Q_ab=cFlow(label='Qab', bus=b_abwaerme_nt, costsPerFlowHour=FuelcostsAbw)
                                      )
            AbwaermeWPs.append(aAbwaermeWP)

        energy_system.addComponents(*AbwaermeWPs)
    # </editor-fold>
    # <editor-fold desc="Speicher">
    if "Speicher" in erz_daten:
        Storages = list()
        for item in erz_daten["Speicher"]:
            # relative Capacity of Storage
            capacity_max_rel = calculate_relative_capacity_of_storage(item, zeitreihen)

            # Investment of Capacity
            Invest_capacity = get_invest_from_excel(item, e_costs, e_funding, years, is_storage=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            capacity = handle_nom_val(item.get("capacity", None))
            nominal_val_flows = handle_nom_val(item.get("nominal_val", None))

            # <editor-fold desc="Investment of Flows">
            # Same Invest Args for both flows. Same Power for in and out flow!
            Invest_flow_in = get_invest_from_excel(item, e_costs, e_funding, years, is_flow_of_storage=True)
            Invest_flow_out = get_invest_from_excel(item, e_costs, e_funding, years, is_flow_of_storage=True)

            help_invest = cEffectType(label=f"helpInv{item['label']}", unit="", description="InvestHelp",
                                      min_investSum=0, max_investSum=0)
            energy_system.addEffects(help_invest)
            Invest_flow_in.specificCosts[help_invest] = -1
            Invest_flow_out.specificCosts[help_invest] = 1
            # </editor-fold>

            aStorage = cStorage(label="Speicher" + item["label"],
                                exists=exists, group=item.get("group", None),
                                inFlow=cFlow(label="QthLoad", bus=b_fernwaerme,
                                             nominal_val=nominal_val_flows, exists=exists, investArgs=Invest_flow_in),
                                outFlow=cFlow(label="QthUnload", bus=b_fernwaerme,
                                              nominal_val=nominal_val_flows, exists=exists, investArgs=Invest_flow_out),
                                avoidInAndOutAtOnce=True,
                                max_rel_chargeState=capacity_max_rel, fracLossPerHour=item["fracLossPerHour"],
                                capacity_inFlowHours=capacity,
                                investArgs=Invest_capacity,
                                # chargeState0_inFlowHours='lastValueOfSim',
                                # charge_state_end_min=min_capacity*0.5,
                                # chargeState0_inFlowHours=min_capacity*0.5,
                                # charge_state_end_min=min_capacity*0.5,
                                )
            Storages.append(aStorage)
        energy_system.addComponents(*Storages)
    # </editor-fold>
    # <editor-fold desc="Rueckkuehler">
    if "Rueckkuehler" in erz_daten:
        Coolers = list()
        for item in erz_daten["Rueckkuehler"]:
            BusInput = b_strom_bezug
            Fuelcosts = {e_costs: preiszeitreihen["costsBezugEl"] + item.get("ZusatzkostenEnergieInput", 0)}

            # Beschränkung Tageszeit: Nur STunde 8-20
            if item['Beschränkung Einsatzzeit']:
                max_rel = np.tile(np.concatenate((np.zeros(8), np.ones(12), np.zeros(4))), int(len(a_time_series) / 24))
            else:
                max_rel = 1

            # Investment
            invest = get_invest_from_excel(item, e_costs, e_funding, years, is_flow=True)
            # Existenz Zeitreihe
            exists = handle_operation_years(item, years)

            nominal_val = handle_nom_val(item.get("nominal_val", None))

            aCooler = cCoolingTower(label="Rueckkuehler" + item["label"],
                                    specificElectricityDemand=item.get("specificElectricityDemand", None),
                                    exists=exists, group=item.get("group", None),
                                    Q_th=cFlow(label='Qth', bus=b_fernwaerme, exists=exists, max_rel=max_rel,
                                               nominal_val=nominal_val,
                                               costsPerRunningHour={e_costs: item.get("costsPerRunningHour", None)},
                                               investArgs=invest),
                                    P_el=cFlow(label='Pel', bus=BusInput, costsPerFlowHour=Fuelcosts)
                                    )
            Coolers.append(aCooler)
        energy_system.addComponents(*Coolers)
    # </editor-fold>

    from flixOpt_excel.Evaluation.flixPostprocessingXL import cModelVisualizer, cVisuData
    visu_data = cVisuData(es=energy_system)
    model_visualization = cModelVisualizer(visu_data)
    model_visualization.Figure.show()

    # <editor-fold desc="HilfsComp für Preiszeitreihen">
    Pout1 = cFlow(label="Strompreis", bus=b_strom_bezug, nominal_val=0,
                  costsPerFlowHour={e_costs: preiszeitreihen["costsBezugEl"]})
    Pout2 = cFlow(label="Gaspreis", bus=b_gas, nominal_val=0, costsPerFlowHour={e_costs: preiszeitreihen["costsBezugGas"]})
    Pout3 = cFlow(label="Wasserstoffpreis", bus=b_wasserstoff, nominal_val=0,
                  costsPerFlowHour={e_costs: preiszeitreihen["costsBezugH2"]})
    Pout4 = cFlow(label="EBSPreis", bus=b_ebs, nominal_val=0, costsPerFlowHour={e_costs: preiszeitreihen["costsBezugEBS"]})

    Comp = cBaseLinearTransformer(label="HelperPreise",
                                  inputs=[], outputs=[Pout1, Pout2, Pout3, Pout4],
                                  factor_Sets=[{Pout1: 1, Pout2: 1, Pout3: 1, Pout4: 1}]
                                  )
    energy_system.addComponents(Comp)
    # </editor-fold>
    def file_management():
        pass

    # <editor-fold desc="Directory and FileManagement">
    path_excel_templates = os.path.join(os.getcwd(),
                                      "flixOpt_excel/resources/ExcelTemplates")  # Directory of the excelTemplate

    # Check if directory in Final Destination exists yet, else add an index to The Name
    for i in range(0, 50):
        if i == 0:
            new_calc_name = calc_name
        else:
            new_calc_name = calc_name + "_" + str(i)

        pathDir = os.path.join(results_directory, new_calc_name)  # Directory for Final Results
        if os.path.exists(pathDir):
            pass
        else:
            os.makedirs(pathDir)
            break
        if i == 49:  # Limit number of directories to 10
            raise Exception(
                "Maximum number of Directories with this name. Please Change name of calulation")

    results_directory = pathDir
    rel_path_calc_results = os.path.join("./_temp_results",
                                      str(random.randint(10000, 99999)))  # Temporary Directory for Results
    path_results_temp = os.path.join(os.getcwd(),
                                     rel_path_calc_results)  # This is the directory to load the results from later on for saving in the specified directory
    results_directory_used_input = os.path.join(results_directory,
                                             "UsedInputData")  # This is the directory the used Input Data is stored to

    # Copy the input excel files (Erzeuger and Input Data) into the results folder, for later comparison or reuse
    os.makedirs(results_directory_used_input)
    shutil.copy2(excel_file_path, os.path.join(results_directory_used_input, "DataInput.xlsx"))

    # <editor-fold desc="Preprocessing Results to excel">
    with pd.ExcelWriter(os.path.join(results_directory_used_input, "PreProcessing.xlsx"), mode="w",
                        engine="openpyxl") as writer:
        zeitreihen.to_excel(writer, index=True, sheet_name="Zeitreihen")
    # </editor-fold>

    # </editor-fold>
    # <editor-fold desc="Modeling and Solving">
    # choose used timeindexe:
    chosen_es_time_indexe = None  # all timeindexe are used
    # chosen_es_time_indexe = [1,2,3,4,5,6,7]

    aCalc = cCalculation(calc_name, energy_system, 'pyomo', pathForSaving=rel_path_calc_results,
                         chosenEsTimeIndexe=chosen_es_time_indexe)  # create Calculation
    aCalc.doModelingAsOneSegment()  # mathematic modeling of system

    # String-resources
    print("###############################################")
    print("Initiated Comps:")
    for type, comps in erz_daten.items():
        print(f"{type}: {[comp.get('label', None) for comp in comps]}")
    # es.printModel()  # string-output:network structure of model
    # es.printVariables()  # string output: variables of model
    # es.printEquations()  # string-output: equations of model

    # Ausführen der Berechnung
    aCalc.solve(solver_props, nameSuffix='_' + solver_name, aPath=rel_path_calc_results)  # nameSuffix for the results
    # </editor-fold>
    # <editor-fold desc="Saving Solve Results">
    solve_results_path = os.path.join(results_directory, "SolveResults")
    shutil.copytree(path_results_temp, solve_results_path)
    shutil.copy2(excel_file_path, os.path.join(results_directory_used_input, "DataInput.xlsx"))
    # </editor-fold>
    calc_info = str(", ".join([f"pathAnalysis='{results_directory}'",  # Ordner für Analyse
                               f"nameOfCalc='{aCalc.nameOfCalc}'",
                               f"results_folder='{solve_results_path}'",
                               f"outputYears='{years}'"]))

    with open(os.path.join(results_directory, "calc_info.txt"), "w") as log_file:
        log_file.write(calc_info)

    # <editor-fold desc="Post-Processing">
    # Start Postprocessing
    from flixOpt_excel.Evaluation.HelperFcts_post import flixPostXL
    calc1 = flixPostXL(aCalc.nameOfCalc, results_folder=solve_results_path, outputYears=years)

    # <editor-fold desc="Print Main Results">
    MainOutput = calc1.infos["modboxes"]["info"][0]["main_results"]
    with open(os.path.join(calc1.folder, "MainResults.txt"), "w") as log_file:
        pp(MainOutput, log_file)
    # </editor-fold>
    # Write HTML of Model structure to file
    model_visualization.Figure.write_html(os.path.join(calc1.folder, 'Model_structure.html'))

    from flixOpt_excel.Evaluation.graphics_excel import run_excel_graphics_main, run_excel_graphics_years
    run_excel_graphics_main(calc1)
    run_excel_graphics_years(calc1)

    # </editor-fold>

    print(f"Directory for temporary results can be deleted: '{os.path.join(os.getcwd(), rel_path_calc_results)}'")
    print(f" Results are saved under: '{results_directory}'")
