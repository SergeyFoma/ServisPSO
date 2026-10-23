    matched = [
        item[0]
        for item in data_tr
        if item and item[0] and pattern_320.search(item[0])
    ]
    print('matched: ', matched, len(matched))