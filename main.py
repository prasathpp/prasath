from turtle import Screen
from snake import Snake
from food import Food
from scorecard import ScoreCard
import time

screen = Screen()
screen.setup(600,600)
screen.bgcolor("black")
screen.title("My Snake Game")
screen.tracer()

snake = Snake()
food = Food()
score_card = ScoreCard()

game_is_on = True
while game_is_on:
    screen.update()
    time.sleep(0.1)
    snake.move()
    # Detection of collision of food
    if snake.head.distance(food) < 15:
        food.refresh()
        snake.extend()
        score_card.increase_score()
    # Detection of collision of wall
    if snake.head.xcor() <= -290 or snake.head.xcor() >= 290 or snake.head.ycor() <= -290 or snake.head.ycor() >= 290:
        game_is_on = False
        score_card.game_over()
    # Detection of collision of tail
    for segment in snake.segments[1:]:
        if snake.head.distance(segment) < 10:
            game_is_on = False
            score_card.game_over()

screen.exitonclick()
