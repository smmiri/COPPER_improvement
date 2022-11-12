'''create OSW file
'''
def print_to_osw(Htype:str, Scenario:str, HeatingMeter:str):
    file = open("WorkFlow.osw", 'w')
    
    if "old" in Htype:

      file_contents = f'''{{
        "seed_file": "{Htype}.osm",
        "weather_file": "CAN_AB_Calgary.718770_CWEC.epw",
        "steps": [
          {{
            "measure_dir_name": "measures/Infiltration",
            "arguments": {{
            "ach_or_cfm": "ACH (Air Change per hour)",
            "i_val": "8.947",
            "pressure": "@50 Pa",
            "net_oa_flow": "0"
            }}
          }},
          {{
            "measure_dir_name": "measures/Set Exterior Wall Assembly to User Input U Value",
            "arguments": {{
            "u_value": "0.109984",
            "construction": "ASHRAE 189.1-2009 ExtWall Mass ClimateZone 5"
            }}
          }},
          {{
            "measure_dir_name": "measures/Increase R-value of Insulation for Roofs to a Specific Value",
            "arguments": {{
            "r_value": "17.7992"
            }}
          }},        
          {{
            "measure_dir_name": "measures/DR_Measure_setpoint_{Scenario.split('_')[0]}",
            "arguments": {{}}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "Cooling:Electricity",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "{HeatingMeter}",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "Cooling:Electricity",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }},
          
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "{HeatingMeter}",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }}
        ]
      }}
      '''
 
    elif "Apartment" in Htype:

      file_contents = f'''{{
        "seed_file": "{Htype}.osm",
        "weather_file": "CAN_AB_Calgary.718770_CWEC.epw",
        "steps": [
          {{
            "measure_dir_name": "measures/Infiltration",
            "arguments": {{
            "ach_or_cfm": "ACH (Air Change per hour)",
            "i_val": "9.3365",
            "pressure": "@50 Pa",
            "net_oa_flow": "0"
            }}
          }},
          {{
            "measure_dir_name": "measures/Set Exterior Wall Assembly to User Input U Value",
            "arguments": {{
            "u_value": "0.115774",
            "construction": "ASHRAE 189.1-2009 ExtWall Mass ClimateZone 5"
            }}
          }},
          {{
            "measure_dir_name": "measures/Increase R-value of Insulation for Roofs to a Specific Value",
            "arguments": {{
            "r_value": "16.721"
            }}
          }},        
          {{
            "measure_dir_name": "measures/DR_Measure_setpoint_{Scenario.split('_')[0]}",
            "arguments": {{}}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "Cooling:Electricity",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "{HeatingMeter}",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "Cooling:Electricity",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }},
          
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "{HeatingMeter}",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }}
        ]
      }}
      '''      
    elif "new" in Htype:

      file_contents = f'''{{
        "seed_file": "{Htype}.osm",
        "weather_file": "CAN_AB_Calgary.718770_CWEC.epw",
        "steps": [
          {{
            "measure_dir_name": "measures/Infiltration",
            "arguments": {{
            "ach_or_cfm": "ACH (Air Change per hour)",
            "i_val": "5.831",
            "pressure": "@50 Pa",
            "net_oa_flow": "0"
            }}
          }},
          {{
            "measure_dir_name": "measures/Set Exterior Wall Assembly to User Input U Value",
            "arguments": {{
            "u_value": "0.072303",
            "construction": "ASHRAE 189.1-2009 ExtWall Mass ClimateZone 5"
            }}
          }},
          {{
            "measure_dir_name": "measures/Increase R-value of Insulation for Roofs to a Specific Value",
            "arguments": {{
            "r_value": "30.5383"
            }}
          }},        
          {{
            "measure_dir_name": "measures/DR_Measure_setpoint_{Scenario.split('_')[0]}",
            "arguments": {{}}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "Cooling:Electricity",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/AddMeter",
            "arguments": {{
            "meter_name": "{HeatingMeter}",
            "reporting_frequency":"hourly"
            }}
          }},
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "Cooling:Electricity",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }},
          
          {{
            "measure_dir_name": "measures/export_meterto_csv_custom_folder",
            "arguments": {{
              "meter_name": "{HeatingMeter}",
              "reporting_frequency":"Hourly",
              "sim_label":"{Htype}_{Scenario}"
            }}
          }}
        ]
      }}
      '''        

    #print to file
    file.write(file_contents)
      
    #close file
    file.close()

'''run from command line
'''
def run_from_cl(PaTh:str):
    import os
    OG_directory = 'C:/Users/smoha/Documents/archetypes_base'
    os.chdir("C:/openstudio-3.4.0/bin")
    os.system(F'openstudio.exe run --workflow C:/Users/smoha/Documents/archetypes_base/{PaTh}.osw')
    os.chdir(OG_directory)