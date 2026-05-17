"""Image handling module for documents."""
import os
import base64
import requests
from io import BytesIO
from PIL import Image

class ImageHandler:
    def __init__(self, cache_dir: str = "outputs/cache"):
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)
        
    def process_image(self, source: str, max_size: tuple[int, int] = (1920, 1080)) -> str:
        """Process an image from URL, base64, or local path and return path to cached processed file."""
        if source.startswith("http://") or source.startswith("https://"):
            return self._from_url(source, max_size)
        elif source.startswith("data:image"):
            return self._from_base64(source, max_size)
        elif os.path.exists(source):
            return self._process_pil(Image.open(source), max_size)
        else:
            raise ValueError(f"Invalid image source: {source[:30]}...")

    def _from_url(self, url: str, max_size: tuple) -> str:
        import hashlib
        url_hash = hashlib.md5(url.encode()).hexdigest()
        cached_path = os.path.join(self.cache_dir, f"{url_hash}.png")
        if os.path.exists(cached_path):
            return cached_path
            
        try:
            resp = requests.get(url, timeout=10)
            resp.raise_for_status()
            img = Image.open(BytesIO(resp.content))
            return self._process_pil(img, max_size, cached_path)
        except Exception as e:
            raise ValueError(f"Failed to fetch image from URL: {str(e)}")
            
    def _from_base64(self, data: str, max_size: tuple) -> str:
        import hashlib
        try:
            # Handle standard base64 prefix
            if "," in data:
                data = data.split(",")[1]
            img_data = base64.b64decode(data)
            
            data_hash = hashlib.md5(img_data).hexdigest()
            cached_path = os.path.join(self.cache_dir, f"{data_hash}.png")
            if os.path.exists(cached_path):
                return cached_path
                
            img = Image.open(BytesIO(img_data))
            return self._process_pil(img, max_size, cached_path)
        except Exception as e:
            raise ValueError(f"Failed to decode Base64 image: {str(e)}")
            
    def _process_pil(self, img: Image.Image, max_size: tuple, out_path: str = None) -> str:
        import uuid
        if not out_path:
            out_path = os.path.join(self.cache_dir, f"{uuid.uuid4().hex}.png")
            
        # Convert to RGB if necessary (to save as PNG)
        if img.mode in ('RGBA', 'P', 'LA'):
            img = img.convert('RGBA') # keep alpha for PNG
        else:
            img = img.convert('RGB')
            
        img.thumbnail(max_size, Image.Resampling.LANCZOS)
        img.save(out_path, format="PNG", optimize=True)
        return out_path
