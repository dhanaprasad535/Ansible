def read_from_excel(testcase_failmsg_dict):
    testcase_count = {}

    list1 = []
    for testcase, message in testcase_failmsg_dict.items():
        if message not in list1:
            ''' add the test case and message if it doesn't exist in list1 '''
            message = message.replace(",", '')
            list1.append(message)
            lastchar_3 = testcase[-3:]

            if lastchar_3.isdigit():
                testcase = testcase[:len(testcase) - 3]

            tc_error_list = [testcase, message]

            a = str(tc_error_list)
            testcase_count[a] = 1
        else:
            ''' Increment the value of error and add the testcase to the key '''
            tc_error_list = [testcase, message]
            a = list(tc_error_list)
            ''' loop through keys(testcase, error) and value(count)'''
            for index, kv in enumerate(list(testcase_count.items())):
                ''' To get only error message from key value pair '''
                error_message = [list(kv)[0].split(",")[-1].strip("]")]
                list_new = []
                x = ''
                ''' To remove end spaces and new line char from the strings in list2'''
                for i in error_message:
                    for k in i.split(" "):
                        z = k.strip().replace('\\n', '')
                        list_new.append(z)
                        x = ' '.join(list_new)
                    print(x)
                    x = x.replace("'", '')
                    a[1] = a[1].strip().replace("\n", '')
                    ''' Compare testcase_count items with testcase_failmsg_dict '''
                    x = x.replace('\\', '').replace("'", '').replace('"', '')
                    a[1] = a[1].replace('\\', '').replace("'", '').replace('"', '')
                    print(f"Value of x is {x.strip()}")
                    print(f"Value of a[1] is {a[1].strip()}")
                    if x.strip() == a[1].strip():
                        print("Matched")
                        ''' Get the index position of matched error message using key '''
                        pos = list(testcase_count.keys()).index(kv[0])
                        for ind, keys in enumerate(list(testcase_count.keys())):
                            if ind == pos:
                                ''' Get everything except error message from keyvalue pair'''
                                split_word = list(kv)[0].split(",")[:-1]
                                split_word_new = ''
                                ''' Adding new test case to exsiting test case string '''
                                for i in range(0, len(split_word)):
                                    split_word_new = split_word_new + "," + split_word[i].strip("[")
                                last3_char = split_word_new[-4:-1]
                                if last3_char.isdigit():
                                    split_word_new = split_word_new[:(len(split_word_new)-1) - 3] + "'"

                                testcase_key = "[" + split_word_new + ",'" + a[0] + "'" + ",'" + a[1] + "'" + "]"

                                ''' Replace existing key with new key and delete the old key(testcase, error message)'''
                                testcase_count[testcase_key] = testcase_count[list(testcase_count)[pos]]
                                del testcase_count[list(testcase_count)[pos]]
                                ''' Increment the value'''
                                testcase_count[testcase_key] = testcase_count[testcase_key] + 1
                    else:
                        print("Not matched")
    print(testcase_count)
    return testcase_count
