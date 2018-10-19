import re

from textblob import TextBlob

# text ='In  the  Middle  East  countries  including  Iran,  cardiac  diseases  are  turning  into  major  health  and  social  problems.[2]  Despite  technological  developments  in  the  treatment of cardiovascular'
#
# blob = TextBlob(text)
# res = blob.sentences
#
# print(type(str(res[0])))
# print(str(res[0]))
#
# print( blob.sentences)


# totalCount = 'Inclusion criteria: Aged over 18 years, diagnosed MI, CABG surgery, or acute coronarysyndrome (ACS), ' \
#              'and able to perform a maximal treadmill testExclusion criteria: Heart failure, severe arrhythmias, drug ' \
#              'abuse, or a medical conditioncontraindicative to high-intensity trainingN randomised: total: 90;' \
#              ' home-based cardiac rehabilitation: 28; centre-based cardiacrehabilitation (treadmill exercise): 34;' \
#              ' centre-based cardiac rehabilitation (group exercise): 28Method of assessment: NRDiagnosis (% of pts):' \
#              'Previous AMI: home-based cardiac rehabilitation: 71.4%; treadmill exercise: 67.6%group exercise: 64.3%' \
#              'Previous CABG: home-based cardiac rehabilitation: 21.4%; treadmill exercise: 26.5%;group exercise: 25.0%ACS: ' \
#              'home-based cardiac rehabilitation: 7.2%; treadmill exercise: 5.9% group exercise:10.70%Age (mean \u5364 SD): ' \
#              'total: NR; home-based cardiac rehabilitation: 58 \u5364 8 years; treadmillexercise: 56 \u5364 9 years; group exercise: ' \
#              '58 \u5364 8 yearsPercentage male: total: 88.9%; home-based cardiac rehabilitation: 96.4%; treadmillexercise: 82.4%; group exercise: 89.3%Ethnicity: NR'
# totalCount = '12 a 5c '
totalCount = ' In the United Kingdom home-based cardiac rehabilitation with[1] a self-help manual ' \
             'supported by a nurse facilitator [3]is a popular method of rehabilitation [14], and was offered to more than 10,000'
totalCount = re.sub("[[](.*?)[]]", "", totalCount)
