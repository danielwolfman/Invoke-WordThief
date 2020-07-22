# import socket programming library 
import socket 
  
# import thread module 
from _thread import *
import threading 
  
lock = threading.Lock() 
max_msg_size = 1024


# thread function 
def threaded(c): 
    client = c.getpeername()
    print('Connected to :', client[0], client[1]) 

    # get doc name
    filename = c.recv(max_msg_size).decode('utf-8')
    
    with open(f'/root/doc_logger/{filename}', 'ab+') as log_file:
        
        # write head title (date and file name)
        log_file.write(c.recv(max_msg_size))
        
        while True:
            data = c.recv(max_msg_size)
            log_file.write(data)
            log_file.flush()
    
  
def Main(): 
    port = 8888
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    s.bind(('', port)) 
    
    print('socket binded to port', port) 
  
    # put the socket into listening mode 
    s.listen(5) 
    print('socket is listening') 
  
    # a forever loop until client wants to exit 
    while True: 
  
        # establish connection with client 
        c, addr = s.accept() 
  
        # Start a new thread and return its identifier 
        start_new_thread(threaded, (c,)) 
    s.close() 
  
  
if __name__ == '__main__': 
    Main() 