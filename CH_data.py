#The dictionary to hold values read from Excel
value_read = {
        'state':'',
        'business_segment':'',
        'business_type':'',
        'curr_coverage':'',
        'business_ownership':'',
        'business_start_date':'',
        'employees_count':0,
        'annual_payroll':0,
        'annual_gross_sales':0,
        'footage':0,
        'address_line':'',
        'city':'',
        'zip_code':00000
}

#List of all states in US - Their codes and full name
states = {'AK': 'Alaska',        'AL': 'Alabama',        'AR': 'Arkansas',        'AS': 'American Samoa',        'AZ': 'Arizona',
          'CA': 'California',    'CO': 'Colorado',       'CT': 'Connecticut',     'DC': 'District of Columbia',  'DE': 'Delaware',
          'FL': 'Florida',       'GA': 'Georgia',        'GU': 'Guam',            'HI': 'Hawaii',                'IA': 'Iowa',
          'ID': 'Idaho',         'IL': 'Illinois',       'IN': 'Indiana',         'KS': 'Kansas',                'KY': 'Kentucky',
          'LA': 'Louisiana',     'MA': 'Massachusetts',  'MD': 'Maryland',        'ME': 'Maine',                 'MI': 'Michigan',
          'MN': 'Minnesota',     'MO': 'Missouri',       'MP': 'Northern Mariana Islands',                       'MS': 'Mississippi',
          'MT': 'Montana',       'NA': 'National',       'NC': 'North Carolina',  'ND': 'North Dakota',          'NE': 'Nebraska',
          'NH': 'New Hampshire', 'NJ': 'New Jersey',     'NM': 'New Mexico',      'NV': 'Nevada',                'NY': 'New York',
          'OH': 'Ohio',          'OK': 'Oklahoma',       'OR': 'Oregon',          'PA': 'Pennsylvania',          'PR': 'Puerto Rico',
          'RI': 'Rhode Island',  'SC': 'South Carolina', 'SD': 'South Dakota',    'TN': 'Tennessee',             'TX': 'Texas',
          'UT': 'Utah',          'VA': 'Virginia',       'VI': 'Virgin Islands',  'VT': 'Vermont',               'WA': 'Washington',
          'WI': 'Wisconsin',     'WV': 'West Virginia',  'WY': 'Wyoming'
}

#questions --> XPATH for yes and no radio buttons. Question to Q1,Q2..mapping can be seen in Excel
#format 'q1':['xpath for yes','xpath for no']
questions = {
        'q1':['//*[@id="field_for_CH_327"]/div[1]/div[1]/label','//*[@id="field_for_CH_327"]/div[1]/div[2]/label'],
        'q2':['//*[@id="field_for_CH_300"]/div[1]/div[1]/label','//*[@id="field_for_CH_300"]/div[1]/div[2]/label'],
        'q3':['//*[@id="field_for_CH_301"]/div[1]/div[1]/label','//*[@id="field_for_CH_301"]/div[1]/div[2]/label'],
        'q4':['//*[@id="field_for_CH_302"]/div[1]/div[1]/label','//*[@id="field_for_CH_302"]/div[1]/div[2]/label'],
        'q5':['//*[@id="field_for_CH_303"]/div[1]/div[1]/label','//*[@id="field_for_CH_303"]/div[1]/div[2]/label'],
        'q6':['//*[@id="field_for_CH_304"]/div[1]/div[1]/label','//*[@id="field_for_CH_304"]/div[1]/div[2]/label']
}

#//*[@id="field_for_CH_324"]/div[1]/div[2]/label
#//*[@id="field_for_CH_322"]/div[1]/div[1]/label

question_list = {
    'Do you currently have a Business Owners Policy in effect?':['P','No'],
    'Have you filed any insurance claims for this business in the past five years?':['P','No'],
    'Has your Business Owners Policy insurance coverage been cancelled or non-renewed in the past three years for reasons other than nonpayment of premium?':['P','No'],
    'Do you act as a franchisor?':['P','No'],
    'Are there functioning and operational smoke/heat detectors in all units and/or occupancies?':['P','Yes'],
    'Do you sell any products under your name or label?':['P','No'],
    'Do you hold a current U.S. certificate in your area of expertise?':['C','Yes'],
    'Do you provide any physical rehabilitation services?':['C','No'],
    'Do you perform any procedures under anesthesia on premises?':['C','No'],
    'Do you provide any property management services?':['C','No'],
    'Do you organize or host guided tours?':['C','No'],
    'Do you primarily specialize in international travel?':['C','No'],
    'Do you manufacture any industrial artwork or metalwork?':['C','No'],
    'Do you sell any products besides artwork for display?':['C','No'],
    'Do you and all employees or independent contractors hold valid licenses?':['C','Yes'],
    'Do you provide massage or tanning services?':['C','No'],
    'Do you provide nail services (e.g. manicure or pedicure)?':['C','No'],
    'Do you use any commercial cooking equipment (e.g. deep fat fryer)?':['C','No'],
    'Do you own a farm or growing operation?':['C','No'],
    'Do you sell any alcohol or tobacco?':['C','No'],
    'Do you rent or loan bicycles?':['C','No'],
    'Do you sell any used, refurbished, or pre-owned items?':['C','No'],
    'Are your cleaning services limited to residential locations only?':['C','Yes'],
    'Do you provide any handyperson services?':['C','No'],
    'Do you provide any instruction or services specific to the following: acupuncture, aerial yoga, boxing, bungee jumping, cheerleading, equestrian yoga, esthetician, extreme sports, free running, health care provider, martial arts or weapons training, massage therapy, naturopath, nutritionists, osteopaths, parkour instruction, physical therapy, pole dancing, SCUBA, trapeze, wilderness hiking, or wrestling?':['C','No'],
    'Do you offer aerial photography or videography services?':['C','No'],
    'Do you perform any film shoots or production services?':['C','No'],
    'Hired and non-owned auto coverage - ask the following: -  Do you own any vehicles in the name of the business?':['C','No'],
    'Hired and non-owned auto coverage - ask the following: -  Do any employees or owners use their personal car for business use more than 12 times per year?':['C','No'],
    'Hired and non-owned auto coverage - ask the following:-  Do you provide delivery services?':['C','No'],
    'Have you been involved in or do you plan to be involved in any layoffs, mergers, or acquisitions in the past five years or next two years?':['C','No']
    
}