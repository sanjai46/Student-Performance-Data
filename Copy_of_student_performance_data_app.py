{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "provenance": [],
      "authorship_tag": "ABX9TyPKWNQ0IgaPkIEZvh/ST48Y",
      "include_colab_link": true
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "markdown",
      "metadata": {
        "id": "view-in-github",
        "colab_type": "text"
      },
      "source": [
        "<a href=\"https://colab.research.google.com/github/sanjai46/Student-Performance-Data/blob/main/Copy_of_student_performance_data_app.py\" target=\"_parent\"><img src=\"https://colab.research.google.com/assets/colab-badge.svg\" alt=\"Open In Colab\"/></a>"
      ]
    },
    {
      "cell_type": "code",
      "execution_count": null,
      "metadata": {
        "id": "ErbsvRf4RGaS"
      },
      "outputs": [],
      "source": [
        "import pandas as pd\n",
        "import numpy as np\n",
        "df = pd.read_excel('/content/overall IA3 marks.xlsx')"
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "pip install pandas openpyxl"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "VKejEnjPT-IQ",
        "outputId": "848b6904-bd85-41eb-9444-f5e071c82a8a"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "Requirement already satisfied: pandas in /usr/local/lib/python3.12/dist-packages (2.2.2)\n",
            "Requirement already satisfied: openpyxl in /usr/local/lib/python3.12/dist-packages (3.1.5)\n",
            "Requirement already satisfied: numpy>=1.26.0 in /usr/local/lib/python3.12/dist-packages (from pandas) (2.0.2)\n",
            "Requirement already satisfied: python-dateutil>=2.8.2 in /usr/local/lib/python3.12/dist-packages (from pandas) (2.9.0.post0)\n",
            "Requirement already satisfied: pytz>=2020.1 in /usr/local/lib/python3.12/dist-packages (from pandas) (2025.2)\n",
            "Requirement already satisfied: tzdata>=2022.7 in /usr/local/lib/python3.12/dist-packages (from pandas) (2026.1)\n",
            "Requirement already satisfied: et-xmlfile in /usr/local/lib/python3.12/dist-packages (from openpyxl) (2.0.0)\n",
            "Requirement already satisfied: six>=1.5 in /usr/local/lib/python3.12/dist-packages (from python-dateutil>=2.8.2->pandas) (1.17.0)\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "# Select subject columns (DM, PP, CA, OS, DAA, AI)\n",
        "subject_columns = df.columns[3:9]\n",
        "\n",
        "# Convert subject columns to numeric, coercing errors to NaN\n",
        "for col in subject_columns:\n",
        "    df[col] = pd.to_numeric(df[col], errors='coerce')\n",
        "\n",
        "# Calculate statistics for each student\n",
        "df['Average Marks'] = df[subject_columns].mean(axis=1)\n",
        "df['Maximum Marks'] = df[subject_columns].max(axis=1)\n",
        "df['Minimum Marks'] = df[subject_columns].min(axis=1)\n",
        "df['Standard Deviation'] = df[subject_columns].std(axis=1)\n",
        "\n",
        "# Display result\n",
        "print(df[['Average Marks', 'Maximum Marks',\n",
        "          'Minimum Marks', 'Standard Deviation']])\n",
        "\n",
        "# Save to new Excel file\n",
        "output_file = \"student_statistics.xlsx\"\n",
        "df.to_excel(output_file, index=False)\n",
        "\n",
        "print(f\"\\nStatistics saved to {output_file}\")"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "IbQl0yr8T-OQ",
        "outputId": "6e24e906-6ca9-4303-8e9e-dde9b9f8b887"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "      Average Marks  Maximum Marks  Minimum Marks  Standard Deviation\n",
            "0         81.666667           90.0           74.0            6.218253\n",
            "1         58.000000           79.0           24.0           18.973666\n",
            "2         85.333333           96.0           80.0            5.501515\n",
            "3         79.166667           93.0           67.0           10.166940\n",
            "4         70.666667           83.0           50.0           11.093542\n",
            "...             ...            ...            ...                 ...\n",
            "1087      53.500000           66.0           30.0           13.663821\n",
            "1088      69.333333           80.0           53.0           10.013324\n",
            "1089      46.666667           70.0           24.0           18.007406\n",
            "1090      78.000000           90.0           60.0           12.961481\n",
            "1091      66.166667           70.0           57.0            4.792355\n",
            "\n",
            "[1092 rows x 4 columns]\n",
            "\n",
            "Statistics saved to student_statistics.xlsx\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "print(df.describe())\n"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "EW6IKN4EUlOd",
        "outputId": "f6593d56-35c7-4b6d-e4b9-75561d8126e9"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "              S.NO           DM           PP           CA           OS  \\\n",
            "count  1092.000000  1087.000000  1088.000000  1088.000000  1082.000000   \n",
            "mean    546.500000    72.968721    77.641544    74.842831    71.476895   \n",
            "std     315.377552    17.613587    13.482286    12.820828    14.814172   \n",
            "min       1.000000     0.000000     8.000000     2.000000     5.000000   \n",
            "25%     273.750000    62.000000    71.000000    68.000000    63.000000   \n",
            "50%     546.500000    76.000000    81.000000    77.000000    74.000000   \n",
            "75%     819.250000    86.000000    88.000000    84.000000    82.000000   \n",
            "max    1092.000000   100.000000    99.000000    99.000000    99.000000   \n",
            "\n",
            "               DAA          AI      Average  Average Marks  Maximum Marks  \\\n",
            "count  1087.000000  1088.00000  1091.000000    1091.000000    1091.000000   \n",
            "mean     74.620975    75.38511    74.452826      74.452826      85.912007   \n",
            "std      15.566345    14.10317    11.709389      11.709389       9.857496   \n",
            "min       9.000000    10.00000    12.200000      12.200000      19.000000   \n",
            "25%      65.500000    67.00000    68.250000      68.250000      82.000000   \n",
            "50%      78.000000    78.00000    76.833333      76.833333      88.000000   \n",
            "75%      87.000000    86.00000    82.833333      82.833333      92.500000   \n",
            "max      99.000000    99.00000    95.333333      95.333333     100.000000   \n",
            "\n",
            "       Minimum Marks  Standard Deviation  \n",
            "count    1091.000000         1091.000000  \n",
            "mean       60.875344            9.473181  \n",
            "std        15.552959            3.967180  \n",
            "min         0.000000            0.816497  \n",
            "25%        52.000000            6.624450  \n",
            "50%        63.000000            8.961027  \n",
            "75%        72.000000           11.674044  \n",
            "max        91.000000           26.454993  \n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "print(\"overall IA3 marks\")\n",
        "print(df)"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "a4HqPng0VD63",
        "outputId": "df99c5bb-17f5-42da-dda6-33d3bf62c606"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "overall IA3 marks\n",
            "      S.NO                NAME  REG. NO.    DM    PP    CA    OS   DAA    AI  \\\n",
            "0        1     A GAYATHIRIDEVI  24CS0001  83.0  74.0  87.0  80.0  76.0  90.0   \n",
            "1        2           A LOGESSH  24CS0002  24.0  51.0  79.0  67.0  61.0  66.0   \n",
            "2        3           A MOHITHA  24CS0003  80.0  83.0  85.0  84.0  84.0  96.0   \n",
            "3        4         A NEHASHREE  24CS0004  67.0  75.0  87.0  83.0  70.0  93.0   \n",
            "4        5       A RAMKARTHICK  24CS0005  50.0  70.0  73.0  72.0  83.0  76.0   \n",
            "...    ...                 ...       ...   ...   ...   ...   ...   ...   ...   \n",
            "1087  1088           VIJAY T A  24CS3025  45.0  57.0  59.0  66.0  30.0  64.0   \n",
            "1088  1089            YOGESH M  24CS3026  62.0  71.0  80.0  75.0  53.0  75.0   \n",
            "1089  1090       YOGESHWARAN S  24CS3027  24.0  42.0  30.0  52.0  70.0  62.0   \n",
            "1090  1091          SANCHITA V       NaN  60.0  90.0  88.0  70.0  70.0  90.0   \n",
            "1091  1092  MOHAMED ABDULLAH R       NaN  65.0  68.0  57.0  68.0  69.0  70.0   \n",
            "\n",
            "        Average         Category  Average Marks  Maximum Marks  Minimum Marks  \\\n",
            "0     81.666667  Average Student      81.666667           90.0           74.0   \n",
            "1     58.000000  Average Student      58.000000           79.0           24.0   \n",
            "2     85.333333           Topper      85.333333           96.0           80.0   \n",
            "3     79.166667  Average Student      79.166667           93.0           67.0   \n",
            "4     70.666667  Average Student      70.666667           83.0           50.0   \n",
            "...         ...              ...            ...            ...            ...   \n",
            "1087  53.500000  Average Student      53.500000           66.0           30.0   \n",
            "1088  69.333333  Average Student      69.333333           80.0           53.0   \n",
            "1089  46.666667     Slow Learner      46.666667           70.0           24.0   \n",
            "1090  78.000000  Average Student      78.000000           90.0           60.0   \n",
            "1091  66.166667  Average Student      66.166667           70.0           57.0   \n",
            "\n",
            "      Standard Deviation  \n",
            "0               6.218253  \n",
            "1              18.973666  \n",
            "2               5.501515  \n",
            "3              10.166940  \n",
            "4              11.093542  \n",
            "...                  ...  \n",
            "1087           13.663821  \n",
            "1088           10.013324  \n",
            "1089           18.007406  \n",
            "1090           12.961481  \n",
            "1091            4.792355  \n",
            "\n",
            "[1092 rows x 15 columns]\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "print(df.describe())"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "MG9fLkn9VEMx",
        "outputId": "4b0c353b-bced-47d0-83fd-a94a393ebd63"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "              S.NO           DM           PP           CA           OS  \\\n",
            "count  1092.000000  1087.000000  1088.000000  1088.000000  1082.000000   \n",
            "mean    546.500000    72.968721    77.641544    74.842831    71.476895   \n",
            "std     315.377552    17.613587    13.482286    12.820828    14.814172   \n",
            "min       1.000000     0.000000     8.000000     2.000000     5.000000   \n",
            "25%     273.750000    62.000000    71.000000    68.000000    63.000000   \n",
            "50%     546.500000    76.000000    81.000000    77.000000    74.000000   \n",
            "75%     819.250000    86.000000    88.000000    84.000000    82.000000   \n",
            "max    1092.000000   100.000000    99.000000    99.000000    99.000000   \n",
            "\n",
            "               DAA          AI      Average  Average Marks  Maximum Marks  \\\n",
            "count  1087.000000  1088.00000  1091.000000    1091.000000    1091.000000   \n",
            "mean     74.620975    75.38511    74.452826      74.452826      85.912007   \n",
            "std      15.566345    14.10317    11.709389      11.709389       9.857496   \n",
            "min       9.000000    10.00000    12.200000      12.200000      19.000000   \n",
            "25%      65.500000    67.00000    68.250000      68.250000      82.000000   \n",
            "50%      78.000000    78.00000    76.833333      76.833333      88.000000   \n",
            "75%      87.000000    86.00000    82.833333      82.833333      92.500000   \n",
            "max      99.000000    99.00000    95.333333      95.333333     100.000000   \n",
            "\n",
            "       Minimum Marks  Standard Deviation  \n",
            "count    1091.000000         1091.000000  \n",
            "mean       60.875344            9.473181  \n",
            "std        15.552959            3.967180  \n",
            "min         0.000000            0.816497  \n",
            "25%        52.000000            6.624450  \n",
            "50%        63.000000            8.961027  \n",
            "75%        72.000000           11.674044  \n",
            "max        91.000000           26.454993  \n"
          ]
        }
      ]
    }
  ]
}