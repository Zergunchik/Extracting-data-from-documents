# cache_manager.py - Управление кэшем распознанных PDF файлов
import hashlib
import os
import shutil
from pathlib import Path
from typing import Optional, Dict, List
import json
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class PDFCacheManager:
    """Менеджер кэша для распознанных PDF файлов."""
    
    def __init__(self, cache_folder: Optional[Path] = None):
        """
        Инициализирует менеджер кэша.
        
        Args:
            cache_folder: Путь к папке кэша. Если None, используется временная папка системы.
        """
        if cache_folder is None:
            # Используем временную папку в директории скрипта
            script_dir = Path(__file__).parent
            cache_folder = script_dir / ".pdf_cache"
        
        self.cache_folder = cache_folder
        self.cache_folder.mkdir(parents=True, exist_ok=True)
        
        # Файл метаданных для хранения информации о закэшированных файлах
        self.metadata_file = self.cache_folder / "cache_metadata.json"
        self.metadata = self._load_metadata()
        logger.info(f"Менеджер кэша инициализирован. Папка кэша: {self.cache_folder}")
    
    def _load_metadata(self) -> Dict:
        """Загружает метаданные кэша из JSON файла."""
        if self.metadata_file.exists():
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    logger.debug(f"Загружены метаданные кэша: {len(data)} записей")
                    return data
            except Exception as e:
                logger.warning(f"Ошибка при загрузке метаданных: {e}")
                return {}
        logger.debug("Файл метаданных не найден, создаем новый кэш")
        return {}
    
    def _save_metadata(self):
        """Сохраняет метаданные кэша в JSON файл."""
        try:
            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump(self.metadata, f, ensure_ascii=False, indent=2)
            logger.debug(f"Метаданные сохранены: {len(self.metadata)} записей")
        except Exception as e:
            logger.error(f"Ошибка при сохранении метаданных: {e}")
    
    def _calculate_pdf_hash(self, pdf_path: Path) -> str:
        """
        Вычисляет хэш PDF файла для идентификации.
        Использует комбинацию имени файла, размера и времени модификации.
        НЕ использует полный путь, чтобы кэш работал при загрузке из разных папок.
        """
        if not pdf_path.exists():
            logger.warning(f"Файл не найден для вычисления хэша: {pdf_path}")
            return ""
        
        stat = pdf_path.stat()
        # Создаем уникальный ключ на основе имени файла, размера и времени модификации
        # НЕ используем полный путь, чтобы кэш работал при загрузке из разных папок
        key_data = f"{pdf_path.name}:{stat.st_size}:{stat.st_mtime}"
        pdf_hash = hashlib.md5(key_data.encode()).hexdigest()
        logger.debug(f"Вычислен хэш для {pdf_path.name}: {pdf_hash[:8]}...")
        return pdf_hash
    
    def get_cached_excel(self, pdf_path: Path) -> Optional[Path]:
        """
        Проверяет наличие закэшированного XLSX файла для данного PDF.
        
        Args:
            pdf_path: Путь к PDF файлу.
            
        Returns:
            Путь к XLSX файлу если найден, иначе None.
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            logger.warning(f"PDF файл не найден: {pdf_path}")
            return None
        
        logger.info(f"🔍 Проверка кэша для PDF: {pdf_path.name}")
        
        # Вычисляем хэш текущего файла
        current_hash = self._calculate_pdf_hash(pdf_path)
        
        if current_hash in self.metadata:
            cache_info = self.metadata[current_hash]
            cached_excel = self.cache_folder / cache_info['excel_filename']
            
            # Проверяем, существует ли файл
            if cached_excel.exists():
                logger.info(f"   ✅ Кэш найден для {pdf_path.name} → {cached_excel.name}")
                print(f"   ✅ Кэш найден для {pdf_path.name}")
                return cached_excel
            else:
                # Файл кэша удален, удаляем запись из метаданных
                logger.warning(f"   ⚠️ Файл кэша не найден для {pdf_path.name}, удаляем запись")
                print(f"   ⚠️ Файл кэша не найден, удаляем запись")
                self._remove_cache_entry(current_hash)
        
        logger.info(f"   ℹ️ Кэш не найден для {pdf_path.name}, потребуется распознавание")
        print(f"   ℹ️ Кэш не найден для {pdf_path.name}")
        return None
    
    def save_to_cache(self, pdf_path: Path, excel_path: Path) -> Path:
        """
        Сохраняет распознанный XLSX файл в кэш.
        
        Args:
            pdf_path: Путь к исходному PDF файлу.
            excel_path: Путь к XLSX файлу для кэширования.
            
        Returns:
            Путь к закэшированному файлу.
        """
        pdf_path = Path(pdf_path)
        excel_path = Path(excel_path)
        
        if not pdf_path.exists():
            logger.error(f"PDF файл не найден: {pdf_path}")
            return excel_path
        
        if not excel_path.exists():
            logger.error(f"XLSX файл не найден: {excel_path}")
            return excel_path
        
        logger.info(f"💾 Сохранение в кэш: {excel_path.name}")
        
        # Вычисляем хэш на основе актуальных данных файла
        pdf_hash = self._calculate_pdf_hash(pdf_path)
        
        # Создаем имя файла кэша
        cached_filename = f"{pdf_hash}_{pdf_path.stem}_extracted_temp.xlsx"
        cached_path = self.cache_folder / cached_filename
        
        # Перемещаем файл в кэш (вместо копирования) и удаляем оригинал
        try:
            shutil.copy2(excel_path, cached_path)
            # Удаляем временный файл из папки с PDF
            excel_path.unlink()
            logger.debug(f"Временный файл удален: {excel_path}")
            
            # Сохраняем метаданные
            self.metadata[pdf_hash] = {
                'pdf_name': pdf_path.name,
                'pdf_size': pdf_path.stat().st_size,
                'pdf_mtime': pdf_path.stat().st_mtime,
                'pdf_absolute_path': str(pdf_path.absolute()),
                'excel_filename': cached_filename,
                'created_at': datetime.now().isoformat(),
                'original_pdf_path': str(pdf_path.absolute())
            }
            self._save_metadata()
            
            logger.info(f"   💾 Данные успешно сохранены в кэш: {cached_path.name}")
            print(f"   💾 Данные сохранены в кэш: {cached_path.name}")
            
        except Exception as e:
            logger.error(f"⚠️ Ошибка при сохранении в кэш: {e}")
            print(f"⚠️ Ошибка при сохранении в кэш: {e}")
            return excel_path
        
        return cached_path
    
    def _remove_cache_entry(self, pdf_hash: str):
        """Удаляет запись из кэша по хэшу."""
        if pdf_hash in self.metadata:
            cache_info = self.metadata[pdf_hash]
            cached_file = self.cache_folder / cache_info['excel_filename']
            if cached_file.exists():
                try:
                    cached_file.unlink()
                    logger.debug(f"Удален файл кэша: {cached_file}")
                except Exception as e:
                    logger.warning(f"Ошибка при удалении файла кэша: {e}")
            del self.metadata[pdf_hash]
            self._save_metadata()
            logger.debug(f"Запись кэша удалена: {pdf_hash[:8]}...")
    
    def remove_from_cache(self, pdf_path: Path):
        """
        Удаляет кэш для конкретного PDF файла.
        
        Args:
            pdf_path: Путь к PDF файлу.
        """
        pdf_path = Path(pdf_path)
        if pdf_path.exists():
            logger.info(f"Удаление кэша для: {pdf_path.name}")
            pdf_hash = self._calculate_pdf_hash(pdf_path)
            self._remove_cache_entry(pdf_hash)
        else:
            logger.warning(f"Невозможно удалить кэш, файл не найден: {pdf_path}")
    
    def clear_cache(self):
        """Очищает весь кэш."""
        logger.info("🗑️ Начало очистки всего кэша...")
        try:
            # Удаляем все файлы кэша
            deleted_count = 0
            for item in self.cache_folder.iterdir():
                if item.is_file():
                    try:
                        item.unlink()
                        deleted_count += 1
                        logger.debug(f"Удален файл: {item}")
                    except Exception as e:
                        logger.warning(f"Ошибка при удалении файла {item}: {e}")
            
            # Очищаем метаданные
            self.metadata = {}
            if self.metadata_file.exists():
                self.metadata_file.unlink()
            
            logger.info(f"   ✅ Кэш успешно очищен. Удалено файлов: {deleted_count}")
        except Exception as e:
            logger.error(f"⚠️ Ошибка при очистке кэша: {e}")
    
    def get_cache_stats(self) -> Dict:
        """
        Возвращает статистику кэша.
        
        Returns:
            Словарь со статистикой: количество файлов, общий размер.
        """
        total_size = 0
        file_count = 0
        
        for item in self.cache_folder.iterdir():
            if item.is_file() and item.suffix == '.xlsx':
                total_size += item.stat().st_size
                file_count += 1
        
        stats = {
            'file_count': file_count,
            'total_size_bytes': total_size,
            'total_size_mb': round(total_size / (1024 * 1024), 2),
            'cache_folder': str(self.cache_folder)
        }
        logger.info(f"Статистика кэша: {file_count} файлов, {stats['total_size_mb']} MB")
        return stats


# Глобальный экземпляр менеджера кэша
_cache_manager: Optional[PDFCacheManager] = None


def get_cache_manager() -> PDFCacheManager:
    """Возвращает глобальный экземпляр менеджера кэша."""
    global _cache_manager
    if _cache_manager is None:
        _cache_manager = PDFCacheManager()
    return _cache_manager


def init_cache_manager(cache_folder: Optional[Path] = None) -> PDFCacheManager:
    """Инициализирует глобальный менеджер кэша."""
    global _cache_manager
    _cache_manager = PDFCacheManager(cache_folder)
    return _cache_manager


def clear_global_cache():
    """Очищает глобальный кэш."""
    global _cache_manager
    if _cache_manager is None:
        _cache_manager = PDFCacheManager()
    _cache_manager.clear_cache()
