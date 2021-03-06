# coding=utf-8
"""一些生成器方法，生成随机数，手机号，以及连续数字等"""
import random
from faker import Factory


class Generator:

    def __init__(self):
        self.fake = Factory().create('zh_CN')

    def random_phone_number(self):
        """随机手机号"""
        return self.fake.phone_number()

    def random_name(self):
        """随机姓名"""
        return self.fake.name()

    def random_address(self):
        """随机地址"""
        return self.fake.address()

    def random_email(self):
        """随机email"""
        return self.fake.email()

    def random_ipv4(self):
        """随机IPV4地址"""
        return self.fake.ipv4()

    def random_str(self, min_chars=0, max_chars=8):
        """长度在最大值与最小值之间的随机字符串"""
        return self.fake.pystr(min_chars=min_chars, max_chars=max_chars)

    def factory_generate_ids(self, starting_id=1, increment=1):
        """ 返回一个生成器函数，调用这个函数产生生成器，从starting_id开始，步长为increment。 """
        def generate_started_ids():
            val = starting_id
            local_increment = increment
            while True:
                yield val
                val += local_increment
        return generate_started_ids

    def factory_choice_generator(self, values):
        """ 返回一个生成器函数，调用这个函数产生生成器，从给定的list中随机取一项。 """
        def choice_generator():
            my_list = list(values)
            # rand = random.Random()
            while True:
                yield random.choice(my_list)
        return choice_generator


if __name__ == '__main__':

    generator = Generator()
    print(generator.random_phone_number())
    print(generator.random_name())
    print(generator.random_address())
    print(generator.random_email())
    print(generator.random_ipv4())
    print(generator.random_str(min_chars=6, max_chars=8))
    id_gen = generator.factory_generate_ids(starting_id=0, increment=3)()
    for i in range(5):
        print(next(id_gen))

    choices = ['John', 'Sam', 'Lily', 'Rose']
    choice_gen = generator.factory_choice_generator(choices)()
    for i in range(7):
        print(next(choice_gen))
