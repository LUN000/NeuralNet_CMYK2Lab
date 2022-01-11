import pandas as pd
import numpy as np
from win32com.client import Dispatch
from photoshop import Session
import random
from tqdm import tqdm
import time
from comtypes.client import GetActiveObject
import torch
import torch.nn as nn
from torch.utils.data import TensorDataset, DataLoader

print('This is CMYK2Lab based on Fogra51 ICC file','\n')

#Data Loader
class DataSet():
    def __init__(self,Datas):
        self.data = Datas
    def CMYK2Tensor(self):
        Data = self.data
        CMYKlst=[]
        for i in range(len(Data)):
            CMYKv=torch.Tensor([Data.iloc[i][0]/100,Data.iloc[i][1]/100,
                                   Data.iloc[i][2]/100,Data.iloc[i][3]/100])
            CMYKlst.append(CMYKv)
        CMYK=torch.stack(CMYKlst)
        return CMYK
    def Lab2Tensor(self):
        Data = self.data
        Lablst=[]
        for i in range(len(Data)):
            Labv=torch.Tensor([Data.iloc[i][4]/100,Data.iloc[i][5]/128,
                                   Data.iloc[i][6]/128])
            Lablst.append(Labv)
        Lab=torch.stack(Lablst)
        return Lab

path = "C:\python3.7\Scripts\programs\CMYK2Lab_exe" 
Train=pd.read_csv(path+'\FograTrainDataSet.csv',index_col=False)
Test=pd.read_csv(path+'\FograTestDataSet.csv',index_col=False)

TrData, TeData = DataSet(Train), DataSet(Test)
TrCMYK, TeCMYK = TrData.CMYK2Tensor(), TeData.CMYK2Tensor()
TrLab, TeLab = TrData.Lab2Tensor(), TeData.Lab2Tensor()

print(TrCMYK.shape,TeCMYK.shape,TrLab.shape,TeLab.shape)

batch_size = 100
train_set = TensorDataset(TrCMYK,TrLab)
test_set = TensorDataset(TeCMYK,TeLab)
train_loader = DataLoader(train_set, batch_size=batch_size, shuffle=True)
test_loader  = DataLoader(test_set, batch_size=batch_size, shuffle=False)


# model
class MLP(nn.Module):
    def __init__(self):
        super(MLP, self).__init__()
        self.fc1 = nn.Linear(4, 2048)
        self.fc2 = nn.Linear(2048, 512)
        self.fc3 = nn.Linear(512, 256)
        self.fc4 = nn.Linear(256, 32)
        self.fc5 = nn.Linear(32,3)
    def forward(self, x):
        x = self.fc1(x).relu()
        x = self.fc2(x).relu()
        x = self.fc3(x).relu()
        x = self.fc4(x).relu()
        x = self.fc5(x).tanh()
        return x
model = MLP()
print('模型架構:',model)

#Train
epochs = eval(input('訓練回合數:'))
optimizer = torch.optim.Adam(model.parameters())
criterion = nn.MSELoss()
loss_min = 0.005
device = torch.device("cuda" if torch.cuda.is_available() else 'cpu')
print('使用',device,'中')
model.to(device)
model.train()
print(f'開始訓練{epochs}回合')
t1 = time.time()
for epoch in range(epochs):
    loss_total=0
    n=0
    outsave, ysave = [],[]
    for i,data in enumerate(train_loader):
        x, y = data
        x, y = x.to(device), y.to(device)
        optimizer.zero_grad()
        out = model(x)
        # save batchs out&y results in 1 epoch
        out, y = out.to('cpu'), y.to('cpu')
        outlst = out.detach().numpy().tolist()
        ylst = y.detach().numpy().tolist()
        outsave += outlst
        ysave += ylst
        # Loss
        loss = criterion(out, y)
        loss_total = loss_total + loss
        n = n + 1
        # 反向
        loss.backward()
        # 更新
        optimizer.step()
    loss_ave = loss_total / n
    if (epoch % 100 == 0):
        print('epoch = %8d, loss_ave = %20.12f' % (epoch, loss_ave))
    #儲存最佳loss model
    if loss_ave < loss_min:
        torch.save(model.state_dict(), 'bestmodel.pt')
        print(f'save at epoch = {epoch},loss = {loss_ave}')
        loss_min = loss_ave
        outTrsave, yTrsave = outsave, ysave
    #儲存last epoch model
    if epoch == epochs-1:
        torch.save(model.state_dict(), 'model_last.pt')
        print(f'save last epoch = {epoch},loss = {loss_ave}')
t2 = time.time()
Time = t2-t1
print(f'訓練結束，花費{Time}s','\n')
    
def LabRecover(data):
    Recover= []
    try:
        for i in range(len(data)):
            Recover.append([data[i][0]*100,data[i][1]*128,data[i][2]*128])
    except:
        Recover.append([data[0]*100,data[1]*128,data[2]*128])
    return Recover

def CMYKRecover(data):
    Recover= []
    try:
        for i in range(len(data)):
            Recover.append([data[i][0]*100,data[i][1]*100,data[i][2]*100,data[i][3]*100])
    except:
        Recover.append([data[0]*100,data[1]*100,data[2]*100,data[3]*100])
    return Recover

def DeltaE_1976(result,label):
    dE = []
    for i in range(len(result)):
        dE.append(((result[i][0]-label[i][0])**2+(result[i][1]-label[i][1])**2+
                       (result[i][2]-label[i][2])**2)**0.5)
    return dE
#Lab值恢復
outTrRe,yTrRe = LabRecover(outTrsave), LabRecover(yTrsave)
#dE計算
dE_Tr = DeltaE_1976(outTrRe,yTrRe)
dE_TrAll, dE_TrAve = sum(dE_Tr), sum(dE_Tr)/1296
dE_TrMax, dE_TrMin = max(dE_Tr), min(dE_Tr)
MaxE_TrIndex, MinE_TrIndex = dE_Tr.index(dE_TrMax), dE_Tr.index(dE_TrMin)
MaxE_TrColor, MinE_TrColor = yTrRe[MaxE_TrIndex], yTrRe[MinE_TrIndex]
MaxE_TrLabOut,MinE_TrLabOut = outTrRe[MaxE_TrIndex], outTrRe[MinE_TrIndex]

print('平均Delta_E:', dE_TrAve)
print('最大Delta_E:', dE_TrMax,'\n最大Delta_ELab:', MaxE_TrColor,'\n最大Delta_E算出Lab:', MaxE_TrLabOut,)
print('最小Delta_E:', dE_TrMin,'\n最小Delta_ELab:', MinE_TrColor,'\n最小Delta_E算出Lab:', MinE_TrLabOut)

print('開始測試')
#Test
model = MLP()
model.load_state_dict(torch.load('bestmodel.pt'))
model.to(device)
model.eval()

for i,data in enumerate(test_loader):
    xte, yte = data
    xte, yte = xte.to(device), yte.to(device)
    outte = model(xte)

outte, yte = outte.to('cpu'), yte.to('cpu')
outte = outte.detach().numpy().tolist()
yte = yte.detach().numpy().tolist()  


outTeRe,yTeRe = LabRecover(outte), LabRecover(yte)
dE_Te = DeltaE_1976(outTeRe,yTeRe)
dE_TeAll, dE_TeAve = sum(dE_Te), sum(dE_Te)/100
dE_TeMax, dE_TeMin = max(dE_Te), min(dE_Te)
MaxE_TeIndex, MinE_TeIndex = dE_Te.index(dE_TeMax), dE_Te.index(dE_TeMin)
MaxE_TeColor, MinE_TeColor = yTeRe[MaxE_TeIndex], yTeRe[MinE_TeIndex]
MaxE_TeLabOut,MinE_TeLabOut = outTeRe[MaxE_TeIndex], outTeRe[MinE_TeIndex]

print('平均Delta_E:', dE_TeAve)
print('最大Delta_E:', dE_TeMax,'\n最大Delta_E顏色:', MaxE_TeColor,'\n最大Delta_E算出Lab:', MaxE_TeLabOut)
print('最小Delta_E :', dE_TeMin,'\n最小Delta_E顏色:', MinE_TeColor,'\n最小Delta_E算出Lab:', MinE_TeLabOut)


Test2=pd.read_csv(path+'\FograTestDataSet2.csv',index_col=False)
TeData2 = DataSet(Test2)
TeCMYK2 = TeData2.CMYK2Tensor()
TeLab2 = TeData2.Lab2Tensor()
test_set2 = TensorDataset(TeCMYK2,TeLab2)
test_loader2  = DataLoader(test_set2, batch_size=batch_size, shuffle=False)

model = MLP()
model.load_state_dict(torch.load('bestmodel.pt'))
model.to(device)
model.eval()

for i,data in enumerate(test_loader2):
    xte2, yte2 = data
    xte2, yte2 = xte2.to(device), yte2.to(device)
    outte2 = model(xte2)

outte2, yte2 = outte2.to('cpu'), yte2.to('cpu')
outte2 = outte2.detach().numpy().tolist()
yte2 = yte2.detach().numpy().tolist()  

outTeRe2,yTeRe2 = LabRecover(outte2), LabRecover(yte2)
dE_Te2 = DeltaE_1976(outTeRe2,yTeRe2)
dE_TeAll2, dE_TeAve2 = sum(dE_Te2), sum(dE_Te2)/100
dE_TeMax2, dE_TeMin2 = max(dE_Te2), min(dE_Te2)
MaxE_TeIndex2, MinE_TeIndex2 = dE_Te2.index(dE_TeMax2), dE_Te2.index(dE_TeMin2)
MaxE_TeColor2, MinE_TeColor2 = yTeRe2[MaxE_TeIndex2], yTeRe2[MinE_TeIndex2]
MaxE_TeLabOut2, MinE_TeLabOut2 = outTeRe2[MaxE_TeIndex2], outTeRe2[MinE_TeIndex2]

print('平均Delta_E:', dE_TeAve2)
print('最大Delta_E:', dE_TeMax2,'\n最大Delta_E顏色:', MaxE_TeColor2,'\n最大Delta_E算出Lab:', MaxE_TeLabOut2)
print('最小Delta_E :', dE_TeMin2,'\n最小Delta_E顏色:', MinE_TeColor2,'\n最小Delta_E算出Lab:', MinE_TeLabOut2)


#Save file
TrRecord = {'Train': dE_TrAve,
            'Time':Time,
          'TRMaxE':dE_TrMax,
          'TRMaxELab:':MaxE_TrColor,
          'TRMaxE_CalLab:':MaxE_TrLabOut,
          'TRMinE:':dE_TrMin,
          'TRMinELab:':MinE_TrColor,
          'TRMinE_CalLab:':MinE_TrLabOut
           }
TeRecord = {'Test': dE_TeAve,
          'T1MaxE':dE_TeMax,
          'T1MaxELab:':MaxE_TeColor,
          'T1MaxE_CalLab:':MaxE_TeLabOut,
          'T1MinE:':dE_TeMin,
          'T1MinELab:':MinE_TeColor,
          'T1MinE_CalLab:':MinE_TeLabOut}
TeRecord2 = {'Test2': dE_TeAve2,
          'T2MaxE':dE_TeMax2,
          'T2MaxELab:':MaxE_TeColor2,
          'T2MaxE_CalLab:':MaxE_TeLabOut2,
          'T2MinE:':dE_TeMin2,
          'T2MinELab:':MinE_TeColor2,
          'T2MinE_CalLab:':MinE_TeLabOut2}

TrRecord,TeRecord,TeRecord2 = pd.DataFrame(TrRecord),pd.DataFrame(TeRecord),pd.DataFrame(TeRecord2)
Record = TrRecord.join(TeRecord,how='left')
Record = Record.join(TeRecord2,how='left')

Record.to_csv(path+'\CMYK2Lab_Record.csv',mode='a',encoding='utf-8-sig')

