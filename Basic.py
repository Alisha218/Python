print("Hi")
name="Alisha"
age=18
height=5.5
is_std=True


if age>20 :print("Eligable")
else: print("Not")

for i in range(5): print(5)


while(age<20):
    print("not aplicable")
    age=age+1


def greet(name):
    print("Welcome"+name)
    print(f"Hello, {name}!")


greet("Arsh")

num=[1,7,2,9,4]
num.append(6)
print(num[3])



person = {"name": "Alice", "age": 25}
print(person["name"])   



class People:
    def __init__(self,name,age):
        self.name=name
        self.age=age

    def greet(self):
        print(f"My name is {self.name} and my age is {self.age}")
        

person1=People("Asha",6)
person1.greet()



