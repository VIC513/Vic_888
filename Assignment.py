import os
import time
import torch
from torch import nn
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from torch import optim
from torch.utils.data import Dataset, DataLoader
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

# 确保保存目录存在
os.makedirs("CheckPoints", exist_ok=True)
os.makedirs("Results", exist_ok=True)

# 1. 构建深度学习模型 (全连接神经网络 FNN)
class NeuralNetwork(nn.Module):
    def __init__(self, input_dim, hidden_dim, output_dim, num_layers=3, dropout_rate=0.2):
        super(NeuralNetwork, self).__init__()
        layers = []
        current_dim = input_dim
        
        # 构建隐藏层
        for _ in range(num_layers):
            layers.append(nn.Linear(current_dim, hidden_dim))
            layers.append(nn.BatchNorm1d(hidden_dim)) # 加入批归一化加速收敛
            layers.append(nn.ReLU())
            layers.append(nn.Dropout(p=dropout_rate)) # Dropout 防止过拟合
            current_dim = hidden_dim
            hidden_dim = hidden_dim // 2 # 逐层减小隐藏层维度
            
        # 输出层
        layers.append(nn.Linear(current_dim, output_dim))
        self.encoder = nn.Sequential(*layers)
        
    def forward(self, x):
        return self.encoder(x)

# 2. 自定义 PyTorch 数据集
class NIRDataSet(Dataset):
    def __init__(self, data, labels):
        super(NIRDataSet, self).__init__()
        self.data = data
        self.labels = labels

    def __len__(self):
        return self.data.shape[0]

    def __getitem__(self, index):
        # 将数据转换为张量并移动到设备 (根据可用性选择 cuda 或 cpu)
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        data_item = torch.tensor(self.data[index], dtype=torch.float).to(device)
        labels_item = torch.tensor(self.labels[index], dtype=torch.float).to(device)
        return data_item, labels_item

if __name__ == '__main__':
    # 检查设备
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print(f"Using device: {device}")

    # ==========================
    # 1. 数据预处理
    # ==========================
    # 读取数据
    df = pd.read_excel("玉米的近红外光谱数据.xlsx")
    
    # 前4列为成分含量 (水分 Moisture、油脂 Oil、蛋白质 Protein、淀粉 Starch)
    target_names = ['Moisture', 'Oil', 'Protein', 'Starch']
    labels = df.iloc[:, 0:4].values
    
    # 第5列到最后为光谱数据 (700个通道)
    nir_data = df.iloc[:, 4:].values
    wave_numbers = df.columns[4:]
    
    # 划分训练集和测试集 (70%训练, 30%测试)
    X_train, X_test, y_train, y_test = train_test_split(nir_data, labels, test_size=0.3, random_state=42)
    
    # 归一化处理 (使用 Standard Scaler)
    scaler_X = StandardScaler()
    X_train_scaled = scaler_X.fit_transform(X_train)
    X_test_scaled = scaler_X.transform(X_test)
    
    # 可视化原始光谱数据示例 (随机挑选几个样本)
    plt.figure(figsize=(10, 5))
    for i in range(5):
        plt.plot(wave_numbers, nir_data[i, :])
    plt.title("Sample NIR Spectra")
    plt.xlabel("Wavelength (nm)")
    plt.ylabel("Absorbance")
    plt.savefig("Results/nir_spectra_samples.png", dpi=300)
    plt.close()

    # ==========================
    # 2. 模型设计与训练准备
    # ==========================
    input_dim = X_train_scaled.shape[1]  # 700
    output_dim = y_train.shape[1]        # 4
    hidden_dim = 256
    num_layers = 2
    
    # 创建模型
    model = NeuralNetwork(
        input_dim=input_dim, 
        hidden_dim=hidden_dim, 
        output_dim=output_dim, 
        num_layers=num_layers,
        dropout_rate=0.3
    ).to(device)
    
    # 导入数据
    train_dataset = NIRDataSet(X_train_scaled, y_train)
    test_dataset = NIRDataSet(X_test_scaled, y_test)
    
    # 数据迭代器
    batch_size = 16 # 样本数较少(80个)，选择较小的 batch_size
    train_loader = DataLoader(train_dataset, batch_size=batch_size, shuffle=True)
    test_loader = DataLoader(test_dataset, batch_size=batch_size, shuffle=False)
    
    # 损失函数 (MSE)
    criterion = nn.MSELoss()
    
    # 优化器 (Adam，并加入 L2 正则化 weight_decay 防止过拟合)
    learning_rate = 0.005
    optimizer = optim.Adam(model.parameters(), lr=learning_rate, weight_decay=1e-4)
    
    # 学习率调度器
    scheduler = optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode='min', factor=0.5, patience=50)

    # ==========================
    # 3. 模型训练
    # ==========================
    num_epochs = 500
    train_loss_list = []
    val_loss_list = []
    
    print("Start Training...")
    for epoch in range(num_epochs + 1):
        # 训练阶段
        model.train()
        epoch_train_loss = 0
        for batch_x, batch_y in train_loader:
            optimizer.zero_grad()
            pred_y = model(batch_x)
            loss = criterion(pred_y, batch_y)
            loss.backward()
            optimizer.step()
            epoch_train_loss += loss.item() * batch_x.size(0)
        
        avg_train_loss = epoch_train_loss / len(train_dataset)
        train_loss_list.append(avg_train_loss)
        
        # 验证阶段
        model.eval()
        epoch_val_loss = 0
        with torch.no_grad():
            for batch_x, batch_y in test_loader:
                pred_y = model(batch_x)
                loss = criterion(pred_y, batch_y)
                epoch_val_loss += loss.item() * batch_x.size(0)
                
        avg_val_loss = epoch_val_loss / len(test_dataset)
        val_loss_list.append(avg_val_loss)
        
        # 调整学习率
        scheduler.step(avg_val_loss)
        
        if epoch % 50 == 0:
            print(f"Epoch [{epoch}/{num_epochs}] | Train Loss: {avg_train_loss:.4f} | Val Loss: {avg_val_loss:.4f}")
            
        # 每100 Epoch 保存模型
        if epoch > 0 and epoch % 100 == 0:
            checkpoint = {
                'epoch': epoch,
                'model_state': model.state_dict(), 
                'optimizer_state': optimizer.state_dict(),
                'val_loss': avg_val_loss
            }
            torch.save(checkpoint, f"CheckPoints/checkpoint_epoch_{epoch}.pth")

    # ==========================
    # 4. 模型评估与结果保存
    # ==========================
    # 绘制损失曲线
    plt.figure(figsize=(8, 5))
    plt.plot(range(len(train_loss_list)), train_loss_list, label='Train Loss')
    plt.plot(range(len(val_loss_list)), val_loss_list, label='Validation Loss')
    plt.xlabel("Epochs")
    plt.ylabel("MSE Loss")
    plt.title("Training and Validation Loss")
    plt.legend()
    plt.savefig("Results/Loss_Curve.png", dpi=300)
    plt.close()

    # 测试集上的预测
    model.eval()
    all_preds = []
    all_targets = []
    with torch.no_grad():
        for batch_x, batch_y in test_loader:
            preds = model(batch_x)
            all_preds.append(preds.cpu().numpy())
            all_targets.append(batch_y.cpu().numpy())
            
    all_preds = np.vstack(all_preds)
    all_targets = np.vstack(all_targets)
    
    # 绘制预测值与真实值的对比图
    fig, axes = plt.subplots(2, 2, figsize=(12, 10))
    axes = axes.flatten()
    
    for i in range(4):
        ax = axes[i]
        ax.scatter(all_targets[:, i], all_preds[:, i], alpha=0.7)
        
        # 绘制 y=x 理想直线
        min_val = min(all_targets[:, i].min(), all_preds[:, i].min())
        max_val = max(all_targets[:, i].max(), all_preds[:, i].max())
        ax.plot([min_val, max_val], [min_val, max_val], 'r--')
        
        # 计算该成分的 MSE
        mse = np.mean((all_targets[:, i] - all_preds[:, i])**2)
        
        ax.set_title(f"{target_names[i]} (MSE: {mse:.4f})")
        ax.set_xlabel("True Values")
        ax.set_ylabel("Predicted Values")
        
    plt.tight_layout()
    plt.savefig("Results/Prediction_vs_True.png", dpi=300)
    plt.close()

    # 将预测结果保存为 CSV
    results_df = pd.DataFrame()
    for i, name in enumerate(target_names):
        results_df[f'{name}_True'] = all_targets[:, i]
        results_df[f'{name}_Pred'] = all_preds[:, i]
        
    results_df.to_csv("Results/Test_Predictions.csv", index=False)
    print("Training finished! Results saved in 'Results' folder.")
