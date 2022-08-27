# Creating function for calculating factorial.
def factorial(a):
    if a==1:
        return 1 #Case 1
    elif a==0:
        return 1 #Case 2
    else:
        return (a*factorial(a-1)) #Recursive function.

a = int(input('Enter the number: ')) #Taking input from user.
if a<0: #Incase number is negetive.
    print("Factorial for negetive numbers does not exists.")
else:
    ans=factorial(a) #Calling the function.
    print("Factorial of ", a, " is: ", ans ) #Printing the final answer.
