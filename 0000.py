def check_contain_chinese(check_str):
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

if __name__ == "__main__":
    print(check_contain_chinese('i love yiu '))
    print(check_contain_chinese('i love you'))
    print(check_contain_chinese('xx中国'))