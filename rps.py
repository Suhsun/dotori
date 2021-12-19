# 이름
# 날짜
# 가위바위보 게임: 플레이어 vs 컴퓨터

import random

# 컴퓨터가 랜덤으로 선택하도록 함수로 변환합니다.
# 문자를 숫자로
computer = random.randint(1,3)

def convert():
    global computer
    if(computer == 1):
        computer = 'r'
    elif(computer == 2):
        computer == 'p'
    else:
        computer =='s'

# 승리 횟수와 점수를 추적합니다.
def score():
    print("\nScore: 플레이어-", userWins, "컴퓨터-", computerWins)
    if(computerWins == userWins and computerWins > 0):
        print("\동점입니다! 한 번 더 해서 컴퓨터를 이겨 주세요!")
    elif(computerWins > userWins):
        print("\n컴퓨터가 이기고 있어요. 컴퓨터를 이길 때까지 도전!")
    elif(computerWins < userWins):
        print("\n당신이 이기고 있어요! 더 큰 점수차로 이겨 볼까요!")

computerWins = 0
userWins = 0

# 게임방법을 출력합니다.
print("가위바위보 게임에 오신 것을 환영합니다!")
print("컴퓨터와 게임을 하려면 다음 규칙에 따라야 합니다.:")
print("'r'을 입력하면 바위\n'p'를 입력하면 보\n's'를 입력하면 가위\n게임 종료를 원하시면 'e'를 선택해 주세요:")
user = input()

# 컴퓨터의 선택 항목 생성
# computer = random.randint(1,3)

while(True):
    # 플레이어가 게임을 종료할 것인지 확인
    if(user == "e"):
        print("게임이 종료되었습니다. 나중에 다시 게임해 주세요!")
        break

# 컴퓨터의 랜덤 숫자를 문자로 변환하는 함수 호출

convert()

# 승자 결정
if(computer == user):
    print("\n동점입니다!")
elif(computer == "r" and user == "s"):
    computerWins = computerWins + 1
    print("\n바위가 가위를 이겼습니다... 컴퓨터가 이겼습니다!")
elif(computer == "r" and user == "p"):
    userWins = userWins + 1
    print("\n보가 바위를 이겼습니다... 당신이 이겼습니다!")
elif(computer == "p" and user == "s"):
    userWins = userWins + 1
    print("\n가위가 보를 이겼습니다... 당신이 이겼습니다!")
elif(computer == "s" and user == "r"):
    userWins = userWins + 1
    print("\n바위가 가위를 이겼습니다... 당신이 이겼습니다!")
elif(computer == "p" and user == "r"):
    computerWins = computerWins + 1
    print("\n보가 바위를 이겼습니다... 컴퓨터가 이겼습니다!")
elif(computer == "s" and user == "p"):
    computerWins = computerWins + 1
    print("\n가위가 보를 이겼습니다... 컴퓨터가 이겼습니다!")
else:
    print("\n잘못된 입력입니다. 'r', 'p', 's', 'e' 중 하나로 선택해 주세요.")

# 점수 출력
score()

# 한판 더 할지 묻기
print("'r'을 입력하면 바위\n'p'를 입력하면 보\n's'를 입력하면 가위:\n게임 종료를 원하시면 'e'를 선택해 주세요:")
user = input()
computer = random.randint(1,3)