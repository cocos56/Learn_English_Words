# print(ord(' '))

# s = '1234'
# s = s[0:3]
# print(s)

# print(len(s))

def remove_end_space(s):
    length = len(s)
    space_ords = [32, 160] #空格有32的空格和160的空格，32的空格为常见空格
    
    while(ord(s[length-1]) in space_ords):
        length -= 1
    
    s = s[:length]
    return s

# s = '2  2'
# s = remove_end_space(s)
# print(s)

# print(len(s))