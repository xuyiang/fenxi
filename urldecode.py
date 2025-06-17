from urllib.parse import unquote, urlparse, parse_qs, urljoin
#自动解码用64位加密的网址
def extract_pdf_url(viewer_url):
    
    try:
        
        parsed = urlparse(viewer_url)
        query = parse_qs(parsed.query)
        
        
        if 'file' not in query:
            raise ValueError("URL 中缺少 file 参数")
        
        encoded_path = query['file'][0]
        
        # 解码路径
        decoded_path = unquote(encoded_path)
        
        # 构建基础 URL
        base_url = f"{parsed.scheme}://{parsed.netloc}"
        
        # 返回完整 PDF URL
        return urljoin(base_url, decoded_path)
    
    except Exception as e:
        print(f"解析失败: {e}")
        return None