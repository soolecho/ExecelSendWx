import logging
import sys
import argparse
import traceback
import signal

from modules.wps_extractor import WPSExtractor
from modules.wechat_sender import WeChatSender
from modules.table_processor import TableProcessor

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("app.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)

for handler in root_logger.handlers:
    handler.setLevel(logging.DEBUG)

logger = logging.getLogger(__name__)

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    
    logger.critical("=" * 80)
    logger.critical("UNHANDLED EXCEPTION!")
    logger.critical("=" * 80)
    logger.critical(f"Exception type: {exc_type.__name__}")
    logger.critical(f"Exception value: {exc_value}")
    logger.critical("Traceback:")
    logger.critical("-" * 80)
    logger.critical("\n".join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
    logger.critical("=" * 80)
    
    print("\n" + "=" * 80)
    print("程序发生未处理的异常！")
    print(f"异常类型: {exc_type.__name__}")
    print(f"异常信息: {exc_value}")
    print("详细日志已记录到 app.log")
    print("=" * 80 + "\n")

def handle_signal(signum, frame):
    logger.critical(f"Received signal: {signum}")
    logger.critical("Traceback at signal:")
    logger.critical("\n".join(traceback.format_stack(frame)))
    print(f"\n收到信号 {signum}，程序即将退出...")
    sys.exit(1)

sys.excepthook = handle_exception

try:
    signal.signal(signal.SIGINT, handle_signal)
    signal.signal(signal.SIGTERM, handle_signal)
except:
    pass


def format_table_data(table_data):
    result = []
    for table in table_data:
        for row in table:
            result.append(" | ".join(row))
        result.append("-" * 50)
    return "\n".join(result)


def send_to_wechat(sender, content, recipient):
    if not content:
        logger.warning(f"No content to send to {recipient}")
        return False
    
    messages = []
    max_message_length = 2000
    for i in range(0, len(content), max_message_length):
        messages.append(content[i:i+max_message_length])
    
    return sender.send_multiple_messages(messages, recipient)


def process_table_filter(config):
    logger.info("Processing table filter mode...")
    
    document_url = config.get("document_url", "")
    if not document_url:
        document_url = input("请输入WPS在线文档URL: ").strip()
    
    sheet_name = config.get("sheet_name", "")
    if not sheet_name:
        sheet_name = input("请输入要操作的Sheet名称(留空则使用第一个): ").strip()
    
    name_column = config.get("name_column", "")
    if not name_column:
        name_column = input("请输入人名所在列的列名(如:姓名): ").strip()
    
    extract_columns = config.get("extract_columns", [])
    if not extract_columns:
        extract_columns_input = input("请输入要提取的列名(用逗号分隔，如:内容,金额): ").strip()
        extract_columns = [col.strip() for col in extract_columns_input.split(",")] if extract_columns_input else []
    
    wechat_column = config.get("wechat_column", "")
    if not wechat_column:
        wechat_column = input("请输入微信昵称所在列的列名(留空则手动指定): ").strip()
    
    logger.info(f"Document URL: {document_url}")
    logger.info(f"Sheet Name: {sheet_name or 'Default'}")
    logger.info(f"Name Column: {name_column}")
    logger.info(f"Extract Columns: {extract_columns}")
    logger.info(f"WeChat Column: {wechat_column or 'Manual'}")
    
    try:
        logger.info("Extracting table data from WPS document...")
        table_data = WPSExtractor.extract_from_document(
            document_url,
            extraction_type="table",
            sheet_name=sheet_name if sheet_name else None
        )
        
        if not table_data:
            logger.warning("No table data extracted")
            print("未从文档中提取到表格数据")
            return
        
        processor = TableProcessor(table_data)
        
        if not processor.get_headers():
            logger.warning("No headers found in table")
            print("表格中未找到表头")
            return
        
        print("\n表格表头:")
        print(" | ".join(processor.get_headers()))
        
        persons = processor.get_all_persons(name_column)
        if not persons:
            logger.warning(f"No persons found in column '{name_column}'")
            print(f"在列'{name_column}'中未找到任何人名")
            return
        
        print(f"\n找到 {len(persons)} 个人:")
        for i, person in enumerate(persons, 1):
            print(f"{i}. {person}")
        
        person_choice = input("\n请选择要发送的人(输入序号或'all'发送给所有人): ").strip()
        
        if person_choice.lower() == "all":
            target_persons = persons
        elif person_choice.isdigit():
            idx = int(person_choice) - 1
            if 0 <= idx < len(persons):
                target_persons = [persons[idx]]
            else:
                print("无效的序号")
                return
        else:
            target_persons = [person_choice]
        
        wechat_mapping = {}
        if wechat_column:
            wechat_mapping = processor.get_person_to_wechat_mapping(name_column, wechat_column)
        
        sender = WeChatSender()
        if not sender.is_online():
            logger.error("WeChat is not online")
            print("微信未登录，请先登录微信")
            return
        
        success_count = 0
        for person in target_persons:
            person_data = processor.get_person_data(person, name_column, extract_columns)
            
            if not person_data:
                print(f"\n未找到 {person} 的相关数据")
                continue
            
            print(f"\n{person} 的数据:")
            print("=" * 60)
            print(person_data)
            print("=" * 60)
            
            recipient = wechat_mapping.get(person, "")
            if not recipient:
                recipient = input(f"请输入 {person} 的微信昵称: ").strip()
            
            if not recipient:
                print(f"跳过发送给 {person}")
                continue
            
            send_confirm = input(f"是否发送给 {recipient}? (y/n): ").strip().lower()
            if send_confirm != "y":
                print(f"已取消发送给 {recipient}")
                continue
            
            if send_to_wechat(sender, person_data, recipient):
                print(f"成功发送给 {recipient}")
                success_count += 1
            else:
                print(f"发送给 {recipient} 失败")
        
        print(f"\n发送完成！成功 {success_count}/{len(target_persons)}")
        
    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        print(f"发生错误: {e}")


def process_simple_mode(config):
    logger.info("Processing simple mode...")
    
    document_url = config.get("document_url", "")
    if not document_url:
        document_url = input("请输入WPS在线文档URL: ").strip()
    
    extraction_type = config.get("extraction_type", "all")
    keyword = config.get("keyword", "")
    css_selector = config.get("css_selector", "")
    
    recipient = config.get("recipient", "")
    if not recipient:
        recipient = input("请输入微信接收人昵称: ").strip()
    
    logger.info(f"Document URL: {document_url}")
    logger.info(f"Extraction Type: {extraction_type}")
    logger.info(f"Recipient: {recipient}")
    
    try:
        logger.info("Extracting content from WPS document...")
        
        if extraction_type == "keyword" and keyword:
            extracted_content = WPSExtractor.extract_from_document(
                document_url,
                extraction_type="keyword",
                keyword=keyword,
                before_chars=config.get("before_chars", 0),
                after_chars=config.get("after_chars", 100)
            )
        elif extraction_type == "table":
            sheet_name = config.get("sheet_name", "")
            extracted_content = WPSExtractor.extract_from_document(
                document_url,
                extraction_type="table",
                sheet_name=sheet_name if sheet_name else None
            )
        elif extraction_type == "selector" and css_selector:
            extracted_content = WPSExtractor.extract_from_document(
                document_url,
                extraction_type="selector",
                selector=css_selector
            )
        else:
            extracted_content = WPSExtractor.extract_from_document(
                document_url,
                extraction_type="all"
            )
        
        if not extracted_content:
            logger.warning("No content extracted from document")
            print("未从文档中提取到任何内容")
            return
        
        logger.info(f"Extracted content length: {len(str(extracted_content))}")
        
        if isinstance(extracted_content, list):
            if isinstance(extracted_content[0], list):
                formatted_content = format_table_data(extracted_content)
            else:
                formatted_content = "\n\n".join(extracted_content)
        else:
            formatted_content = str(extracted_content)
        
        print("\n提取到的内容:")
        print("=" * 60)
        print(formatted_content)
        print("=" * 60)
        
        send_confirm = input("\n是否发送到微信? (y/n): ").strip().lower()
        if send_confirm != "y":
            logger.info("User cancelled sending")
            print("已取消发送")
            return
        
        logger.info("Sending content to WeChat...")
        sender = WeChatSender()
        
        if not sender.is_online():
            logger.error("WeChat is not online")
            print("微信未登录，请先登录微信")
            return
        
        if send_to_wechat(sender, formatted_content, recipient):
            logger.info("Message sent successfully")
            print("\n消息发送成功")
        else:
            logger.error("Failed to send message")
            print("\n发送消息失败")
            
    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        print(f"发生错误: {e}")


def main():
    parser = argparse.ArgumentParser(description="WPS文档提取与微信发送工具")
    parser.add_argument("--no-gui", action="store_true", help="使用命令行模式")
    parser.add_argument("--mode", choices=["simple", "table_filter"], default="simple", help="运行模式")
    args = parser.parse_args()
    
    logger.info("Starting WPS to WeChat automation...")
    
    if not args.no_gui:
        try:
            from modules.gui import run_gui
            run_gui()
            return
        except ImportError as e:
            logger.error(f"无法启动图形界面: {e}")
            print("请安装 PyQt6: pip install pyqt6")
            print("或使用命令行模式: python main.py --no-gui")
            return
    
    mode = args.mode
    
    if mode == "table_filter":
        process_table_filter({})
    else:
        process_simple_mode({})


if __name__ == "__main__":
    main()