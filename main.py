# Pong Game

from turtle import Turtle, Screen
from paddle import Paddle
from ball import Ball
import time

screen = Screen()
screen.setup(height=600, width=800)
screen.bgcolor("black")
screen.title("Pong")
screen.tracer(0)

l_paddle = Paddle((-350,0))
r_paddle = Paddle((350,0))
ball = Ball()
game_is_on = True

r_score = 0
l_score = 0

while game_is_on:
    
    time.sleep(0.07)
    ball.move_ball()
    screen.update()
    # Detect collision with the wall
    if ball.ycor() > 270 or ball.ycor() < -270:
        ball.bounce_y()
    # Detect collision with the paddle
    if ball.distance(r_paddle) < 50 and ball.xcor() > 320 or ball.distance(l_paddle) < 50 and ball.xcor() < -320:
        ball.bounce_x()
    # Detect R paddle misses
    if ball.xcor() > 375:
        ball.reset_position()
        l_score += 1
    # Detect R paddle misses
    if ball.xcor() < -375:
        ball.reset_position()
        r_score += 1

screen.exitonclick()
