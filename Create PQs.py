# author: Peter Kim (Peter.1.Kim@sce.com)
# This script will create PQ schedules by reading 4 excel files, all in the same directory:
# 1. PQName.csv
# 2. PQTime.csv
# 3. PQKW.csv
# 4. PQKVAR.csv
#
# References: '\n' means "new line" and '\t' means "tab"

assetsdir = 'C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/PQ Script/Assets/'   # location of directory of the 4 files mentioned above
outputdir = 'C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/PQ Script/PQs/'  # location of directory of resulting files

names = []
with open(assetsdir + 'PQName.csv', 'r') as f:
    for line in f:
        names.append(line.replace('\n',''))  #names
with open(assetsdir + 'PQTime.csv', 'r') as f:
    for line in f:
        times = line.replace('\n','').split(',')  #times
kw = []
with open(assetsdir + 'PQKW.csv', 'r') as f:
    for line in f:
        kw.append(line.replace('\n','').split(','))  #kw
kvar = []
with open(assetsdir + 'PQKVAR.csv', 'r') as f:
    for line in f:
        kvar.append(line.replace('\n','').split(','))    #kvar

for i in range(len(names)):
    loadname = names[i]
    # writelist = [times[j] + '\t' + str(float(kw[i][j])/1000.0) + '\t' + str(float(kvar[i][j])/1000.0) + '\n' for j in range(len(times))]
    writelist = [str(int(times[j])) + '\t' + str(float(kw[i][j])/1000.0) + '\t' + str(float(kvar[i][j])/1000.0) + '\n' for j in range(len(times))]
    # writelist = [times[j] + '\t' + str(float(kw[i][j])/1000.0) + '\t' + '\n' for j in range(len(times))]
    with open(outputdir + loadname + '.txt', 'w') as writefile:
        writefile.write('2\n')      # write number of columns to be 2
        writefile.write('-1' + '\t' + str(float(kw[i][0])/1000.0) + '\t' + str(float(kvar[i][0])/1000.0) + '\n')    #write inital time stamp -1 to match the same as the first time stamp values of a given schedule
        # writefile.write('-1' + '\t' + str(0) + '\t' + str(0) + '\n')    #write inital time stamp -1 to match the same as the first time stamp values of a given schedule
        # writefile.write('1\n')
        # writefile.write('-1' + '\t' + str(float(kw[i][0])/1000.0) + '\t' + '\n')
        writefile.writelines(writelist)
