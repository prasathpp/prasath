from turtle import Turtle, Screen


screen = Screen()

class Paddle(Turtle):
    def __init__(self, position):
        super().__init__()
        self.shape("square")
        self.color("white")
        self.up()
        self.goto(position)
        self.shapesize(stretch_len=1,stretch_wid=5)
        self.move(position)

    def go_up(self):
        new_y = self.ycor() + 20
        self.goto(self.xcor(), new_y)
    
    def go_down(self):
        new_y = self.ycor() - 20
        self.goto(self.xcor(), new_y)

    def move(self, position):
        screen.listen()
        if position == (-350,0):
            screen.onkeypress(self.go_up, "w")    
            screen.onkeypress(self.go_down, "s")

        elif position == (350,0):
            screen.onkeypress(self.go_up, "Up")    
            screen.onkeypress(self.go_down, "Down")

