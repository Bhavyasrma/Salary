import logging
def bkt_name(key):
    logging.info("Entering in bkt_name function")
    a = key.split('_')
    bkt_name = a[3]
    if a[3] == "1":
        bkt_name = a[3]+"_"+a[2] 
    logging.info("Exiting from bkt_name function")
    return bkt_name 

def get_index_normal(df,v1):
    logging.info("Entering in get_index_normal function")
    logging.debug('DF = \n '+'\t'+ df.to_string().replace('\n', '\n\t'))
    logging.debug("Value = "+str(v1))
    index_list = list(df.columns)
    ind = len(index_list)
    for i,index in enumerate(index_list):
        k = i-1
        to_return_upper = index_list[0]
        if index>v1:
            if k<0:
                k=0
            to_return_upper = index_list[k]
            break
        if v1>index_list[ind-1]:
            to_return_upper = index_list[ind-1]
            break
        elif index==v1:
            to_return_upper = index
            break
        elif v1 == 0:
            to_return_upper = index_list[0]
            break
    logging.debug("to_return_upper = "+str(to_return_upper)) 
    return to_return_upper

def get_index_nm(df,v2):
    logging.info("Entering in get_index_nm function.")
    logging.debug('DF for nm = \n '+'\t'+ df.to_string().replace('\n', '\n\t'))
    logging.debug("Value for nm = "+str(v2))   
    index_list = list(df.index)
    to_return_side = 0
    for i,index in enumerate(index_list):
        k = i-1
        to_return_side = 0
        if index>v2:
            if k<0:
                k=0
            to_return_side = index_list[k]
            break
        elif index==v2:
            to_return_side = index
            break
        elif v2 == 0:
            to_return_side = 0
            break
    logging.debug("to_return_side = "+str(to_return_side))
    return to_return_side  

  