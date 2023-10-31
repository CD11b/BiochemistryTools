# Load module often used for Excel files
from openpyxl import load_workbook

# Load an Excel sheet containing state 1 PDB data
structreData = load_workbook(filename = 'PED00023e003.xlsx')
structreDataSheet = structreData.active

# Constants
atomicWeights = [['H', 1.01], ['C', 12.01], ['N', 14.00], ['O', 16.00], ['S', 32.06]]

# Storing PDB data for each atom in a nested list
atomList = []
rowNumber = 1
for row in structreDataSheet:

    atomList.append([structreDataSheet['G' + str(rowNumber)].value,
                     structreDataSheet['D' + str(rowNumber)].value,
                     structreDataSheet['E' + str(rowNumber)].value,
                     structreDataSheet['F' + str(rowNumber)].value])
    rowNumber = rowNumber + 1

# For the supplied atom ID, use its element ID to figure out its atomic weight based on the constants
def atomMass(atomID):

    global atomicWeight

    for item in atomicWeights:
        if atomID in item[0]:
            atomicWeight = item[1]
            return atomicWeight

# Calculate mass * position vector for given atom
def atomMV(atomID):

    atomMassVector = [atomMass(atom[0]) * atom[1],
                      atomMass(atom[0]) * atom[2],
                      atomMass(atom[0]) * atom[3]]

    return atomMassVector

# Starting values before summation
totalMass = 0
summationMVX = 0
summationMVY = 0
summationMVZ = 0

# Summation of mass * vector term of center of mass formula and summation of atomic weights
for atom in atomList:
    totalMass = totalMass + atomMass(atom[0])
    summationMVX = summationMVX + atomMV(atom)[0]
    summationMVY = summationMVY + atomMV(atom)[1]
    summationMVZ = summationMVZ + atomMV(atom)[2]

# Calculate center of mass using summations
centerOfMass = [summationMVX / totalMass, summationMVY / totalMass, summationMVZ / totalMass]

# Calculate numerator of radius of gyration squared formula for given atom
def atomMVRG(atomID):

    atomMassVectorRG = [(atom[1] - centerOfMass[0])**2,
                        (atom[2] - centerOfMass[1])**2,
                        (atom[3] - centerOfMass[2])**2]

    return atomMassVectorRG

# Staring values for radius of gyration squared summations
summationMVRGX = 0
summationMVRGY = 0
summationMVRGZ = 0

# Summations of numerator of radius of gyration squared
for atom in atomList:
    summationMVRGX = summationMVRGX + atomMVRG(atom)[0]
    summationMVRGY = summationMVRGY + atomMVRG(atom)[1]
    summationMVRGZ = summationMVRGZ + atomMVRG(atom)[2]

# Calculate radius of gyration vector
totalNumberAtoms = len(atomList)
radiusOfGyration = [(summationMVRGX / totalNumberAtoms)**0.5, (summationMVRGY / totalNumberAtoms)**0.5, (summationMVRGZ / totalNumberAtoms)**0.5]

# Calculate magnitude of radius of gyration vector
radiusOfGyrationMagnitude = (radiusOfGyration[0]**2 + radiusOfGyration[1]**2 + radiusOfGyration[2]**2)**0.5
print(radiusOfGyrationMagnitude)
