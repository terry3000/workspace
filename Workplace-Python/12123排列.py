
0


A0 = ['A']
A1 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
A2 = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M',
      'N',  'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
A3 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
A4 = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M',
      'N',  'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
A5 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


for i0 in range(len(A0)):    
    for i1 in range(len(A1)):    
        for i2 in range(len(A2)): 
            for i3 in range(len(A3)):    
                for i4 in range(len(A4)):    
                    # for i5 in range(len(A5)):    
                        print('辽'+A0[i0]+A1[i1]+A2[i2]+A3[i3]+A4[i4],end='')