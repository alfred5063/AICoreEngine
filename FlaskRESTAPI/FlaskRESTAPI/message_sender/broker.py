import pika

# Singleton pattern
class Broker():

  def __init__(self):
    self.connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
    self.channel = self.connection.channel()
    self.channel.queue_declare(queue='hello')

  def send(self, message):
    self.channel.basic_publish(exchange='',
      routing_key='hello',
      body=message
    )
    print('[x] "Process list" sent')
    self.connection.close()
    return(message)
