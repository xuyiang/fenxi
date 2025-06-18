import os
import sys
from pathlib import Path

# 设置环境变量使用镜像源
os.environ['HF_ENDPOINT'] = 'https://hf-mirror.com'

def download_and_save_model():
    """下载并保存模型到本地"""
    try:
        from sentence_transformers import SentenceTransformer
        print("✅ SentenceTransformers 导入成功")
        
        # 创建本地模型目录
        local_model_path = Path("./local_models/all-MiniLM-L6-v2")
        local_model_path.mkdir(parents=True, exist_ok=True)
        
        print(f"🔄 正在下载模型到: {local_model_path.absolute()}")
        
        # 下载模型
        model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')
        print("✅ 模型下载成功")
        
        # 保存到本地
        model.save(str(local_model_path))
        print(f"✅ 模型已保存到本地: {local_model_path.absolute()}")
        
        # 验证本地模型
        print("🔄 验证本地模型...")
        local_model = SentenceTransformer(str(local_model_path))
        
        # 测试编码
        test_sentences = ["这是一个测试句子", "This is a test sentence"]
        embeddings = local_model.encode(test_sentences)
        print(f"✅ 本地模型测试成功，输出形状: {embeddings.shape}")
        
        print(f"\n🎉 模型已成功下载并保存到: {local_model_path.absolute()}")
        print("现在可以离线使用此模型了！")
        
        return str(local_model_path)
        
    except ImportError as e:
        print(f"❌ 导入错误: {e}")
        print("请安装: pip install sentence-transformers")
        return None
        
    except Exception as e:
        print(f"❌ 下载错误: {e}")
        return None

def test_offline_model(model_path):
    """测试离线模型"""
    if not model_path:
        return False
        
    try:
        # 设置离线模式
        os.environ['HF_HUB_OFFLINE'] = '1'
        
        from sentence_transformers import SentenceTransformer
        
        print(f"\n🔄 测试离线模型: {model_path}")
        model = SentenceTransformer(model_path)
        
        # 测试编码
        sentences = ["离线测试句子1", "Offline test sentence 2"]
        embeddings = model.encode(sentences)
        print(f"✅ 离线模型工作正常，输出形状: {embeddings.shape}")
        
        return True
        
    except Exception as e:
        print(f"❌ 离线测试失败: {e}")
        return False

if __name__ == "__main__":
    print("🚀 开始下载 all-MiniLM-L6-v2 模型...")
    
    # 下载并保存模型
    model_path = download_and_save_model()
    
    if model_path:
        # 测试离线模型
        test_offline_model(model_path)
        
        print(f"\n📝 使用说明:")
        print(f"1. 模型保存在: {model_path}")
        print(f"2. 离线使用时，直接加载: SentenceTransformer('{model_path}')")
        print(f"3. 无需网络连接即可使用") 