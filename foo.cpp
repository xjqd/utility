/*
 * =====================================================================================
 *
 *       Filename:  foo.cpp
 *
 *    Description:  
 *
 *        Version:  1.0
 *        Created:  05/06/2015 09:00:36 AM
 *       Revision:  none
 *       Compiler:  gcc
 *
 *         Author:  YOUR NAME (), 
 *   Organization:  
 *
 * =====================================================================================
 */
#include <stdlib.h>
#include <stdio.h>
class Foo
{
    public:
        int public_x;
        int public_y;
    private:
        int private_m;
        int private_n;
    public:
        Foo(int x, int y, int m, int n);
        void set_public_val(int x, int y);
        void set_private_val(int m, int n);
        void set_one_private(int m);
        void print_val();
};
class Foo2 : public Foo
{
    public:
        int public_a;
    private:
        int private_b;
    public:
        Foo2(int x, int y, int m, int n, int a, int b);
        void set_public_val(int x, int a);
        void set_private_val(int m, int b);
        void print_val();
};
Foo::Foo(int x, int y, int m, int n)
{
    public_x = x;
    public_y = y;
    private_m = m;
    private_n = n;
}
void Foo::set_public_val(int x, int y)
{
    public_x = x;
    public_y = y;
    printf("public this = %p\n", this);
}
void Foo::set_private_val(int m, int n)
{
    this->private_m = m;
    this->private_n = n;
    printf("private this = %p\n", this);
}
void Foo::set_one_private(int abc)
{
    this->private_m = abc;
}
void Foo::print_val()
{
    printf("x=%d\n", public_x);
    printf("y=%d\n", public_y);
    printf("this->m=%d\n", this->private_m);
    printf("n=%d\n", private_n);
    printf("this = %p\n", this);
}
//Foo2::Foo2(int x, int y, int m, int n, int a, int b):public_x(x),public_y(y),private_m(m),private_n(n)
Foo2::Foo2(int x, int y, int m, int n, int a, int b):Foo(x,y,m,n)
{
    this->public_a = a;
    this->private_b = b;
}

void Foo2::set_public_val( int x, int a)
{
    public_a = a;
    this->public_a = a;
    x = x;
    this->public_x = x;
    printf("public this = %p\n", this);
}

void Foo2::set_private_val( int m, int b)
{
    printf("private this = %p\n", this);
    //Foo::private_m = m;
    //this->private_m = m;
    //this.set_one_private(m);
    this->set_one_private(m);
    private_b = b;
    this->private_b = b;
}
void Foo2::print_val()
{
    printf("this = %p\n", this);
    printf("x=%d\n", this->public_x);
    printf("y=%d\n", public_y);
    //printf("m=%d\n", this->private_m);
    //printf("n=%d\n", this->private_n);
    printf("a=%d\n", this->public_a);
    printf("b=%d\n", private_b);
}

int main()
{
    Foo foo = Foo(1, 2, 3, 4);
    foo.set_public_val(5,6);
    foo.set_private_val(7, 8);
    foo.public_x = 100;
    //foo.private_m = 100;
    foo.set_one_private(101);
    foo.print_val();
    Foo2 foo2 = Foo2(11, 12, 13, 14, 15, 16);
    foo2.set_public_val(21, 22);
    foo2.set_private_val(31, 32);
    foo2.public_a = 999;
    foo2.public_x = 1000;
    //foo2.private_b = 1001;
    //foo2.private_m = 1002;
    foo2.print_val();
    printf("sizeof(foo)=%d, sizeof(foo2)=%d\n", sizeof(Foo), sizeof(Foo2));
    return 0;
}
