from datetime import datetime
SHEETS = {"E": ("1dQUTRqGoJSPZaHogOd2SlgsjDgFRG3RafuhWyHBYhk4", "Master"),
          "G": ("1QYjdNhCviNur7WwL3RhLrMIfkgm53D30-QVEdSPeQn8", "Master"),
          "Ccopy": ("1zFRQmItbv_8Lu4baQCR2T6lmQyYqEY4m2Dj3q5xtUUQ", "Master"),
          "C": ("1SUtuESSL6DqZ_0q8ykW_vPpVjcgMZpLiDUfPuxjI7JE", "Master"),
          #"B": ("1QYjdNhCviNur7WwL3RhLrMIfkgm53D30-QVEdSPeQn8", "Master"),
          "B": ("1DIWLtcNPzIXrhz32cNvN4-gLOs6U68WfTRXVr08cRxc", "Master"), # COPY
          "D": ("1XmleRCeqgwCgpaJLPzuuCz1yWYnBqF7OBeuKNgeDrK4", "Master"),
          "Q": ("1yA2KUMzQg-4GQRYtBtUQQ3S9Lax4H7E2R8rGQchjXXY", "Master"),
          "N": ("1xmQmD8OrfST2ZOkcmbhbAw8UgXWRWu8Hirohzt8kEno", "Master"),
          "Test": ("1_HMUM6fbhD6P5696wbBpaAcIaLSB2jROfcBAddedjB8", "Master"),
          }
          #"D": ("13Bb_Ir44_pqNlwqiLzFfJW57uA_aQNho3yS6Y9opZ4A", "Master")} # Oxton
          
ROADSHEETS = {"Ccopy": ("1tOuZWe7yoSl7swFAiQ5CFQm1D8Aq-1PYyI1cBMwsxg8", "Addresses"),
              "B": ("1YRcobhngiqHh2kGZFEdht2zmfK9KGxB3RUBWCrNTzDA", "Addresses"),
              "G": ("1oI3a-CfqeHu_HrVWUplBQiYf0yLwlN2knSg3OH_jdvs", "Addresses"),
              "D": ("1abfa6PXvIS0f54L1y7KTi9tdNb7YPU_AcX0DESqzGKE", "Addresses"),
              "Q": ("1OVydwrSVvvBtRSedbVuf69ZgwFQcoBYYGj6AhTWQ6ME", "Addresses"),
              "N": ("1eJyg_gfidbp5MFbx8QBhSN331tZkjHEoq___CjIEGBM", "Addresses"),
              "Test": ("1HJSqGDmo2tZCufy5qRnN0GtLrjW8NmtxYOo1xzJIrA0", "Addresses"),
              }
          
LOCAL_ELECTION = {"C": {"name": "LE23", "last": datetime(2023, 5, 4), "recent": ["MM24"]},
                  "E": {"name": "LE23", "last": datetime(2023, 5, 4), "recent": ["MM24"]},
                  "G": {"name": "LE23", "last": datetime(2023, 5, 4), "recent": ["LE22"]}}
GENERAL_ELECTION = {"C": {"name": "GE24", "last": datetime(2024, 7, 4)},
                    "E": {"name": "GE24", "last": datetime(2024, 7, 4)},
                    "G": {"name": "GE24", "last": datetime(2024, 7, 4)}}
                    
ROADS = {"B": [],
         "C": ["Park Road West"],
         "E": [],
         "G": ["Ffrancon Drive", "Village Road"],
         "Ccopy": ["Park Road West"],
         "D": [],
         "Q": [],
         "Q": [],
         "Test": []
         } # Used where address lines contain a house and a road, but the house does not start with a digit.
