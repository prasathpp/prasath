from turtle import Turtle

SCORECARD_POSITION = (0,270)
ALIGNMENT = "center"
FONT = ("Arial", 14, "bold")

class ScoreCard(Turtle):
    def __init__(self) -> None:
        super().__init__()
        self.up()
        self.hideturtle()
        self.goto(SCORECARD_POSITION)
        self.color("white")
        self.score = 0
        self.update_scoreboard()

    def update_scoreboard(self):
        self.write(f"Score: {self.score}", align=ALIGNMENT, font=FONT)
    
    def game_over(self):
        self.goto(0,0)
        self.write("Game Over", align=ALIGNMENT, font=FONT)
    
    def increase_score(self):
        self.clear()
        self.goto(SCORECARD_POSITION)
        self.score += 1
        self.update_scoreboard()
