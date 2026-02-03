# カードリーダー識別情報確認ツール
# pip install pyscard

from smartcard.System import readers
from smartcard.util import toHexString

def check_readers():
    """接続されているすべてのカードリーダーの識別情報を表示"""
    print("=" * 60)
    print("カードリーダー識別情報")
    print("=" * 60)
    
    try:
        # リーダー一覧取得
        r = readers()
        
        if len(r) == 0:
            print("\n[警告] カードリーダーが見つかりません")
            print("カードリーダーが接続されているか確認してください\n")
            return
        
        print(f"\n検出されたリーダー数: {len(r)}台\n")
        
        # 各リーダーの情報を表示
        for index, reader in enumerate(r):
            print(f"--- リーダー #{index} ---")
            print(f"リーダー名: {reader.name}")
            print(f"リーダーオブジェクト: {reader}")
            print(f"文字列表現: {str(reader)}")
            
            # カードが挿入されているか確認
            try:
                connection = reader.createConnection()
                connection.connect()
                print(f"カード検出: あり")
                
                # カードUIDを取得してみる
                try:
                    get_uid = [0xFF, 0xCA, 0x00, 0x00, 0x00]
                    data, sw1, sw2 = connection.transmit(get_uid)
                    if sw1 == 0x90 and sw2 == 0x00:
                        uid = toHexString(data)
                        print(f"カードUID: {uid}")
                except:
                    print(f"カードUID: 取得失敗")
                
                connection.disconnect()
                
            except Exception as e:
                print(f"カード検出: なし")
            
            print()
        
        print("=" * 60)
        print("\n[推奨される設定方法]")
        print("上記の「リーダー名」を使用して、会議用/授業用を区別できます")
        print("例: config.jsonに以下のように記録")
        print('  "meeting_reader": "リーダー #0 のリーダー名"')
        print('  "class_reader": "リーダー #1 のリーダー名"')
        print()
        
    except Exception as e:
        print(f"\n[エラー] {e}\n")

if __name__ == "__main__":
    check_readers()
    input("\nEnterキーを押して終了...")