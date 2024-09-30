"""
FingerGame 类实现了一个简单的 "压手指" 游戏。
"""

import random
import time
import sys

Finger = str


class FingerGame:
    """
    FingerGame 类实现了一个简单的 "压手指" 游戏。
    """

    def __init__(self) -> None:
        """
        初始化 FingerGame 类。
        """
        self.exit_str = "exit"
        self.available_fingers: list[Finger] = [
            "拇指",
            "食指",
            "中指",
            "无名指",
            "小指",
        ]
        self.win_lose_pairs: list[tuple[Finger, Finger]] = [
            ("拇指", "食指"),
            ("食指", "中指"),
            ("中指", "无名指"),
            ("无名指", "小指"),
            ("小指", "拇指"),
        ]
        self.statistics = self.Statistics(available_fingers=self.available_fingers)

    def __get_input(self) -> str:
        """
        获取用户输入。

        :return: 用户输入的手指或退出指令
        """
        input_prompt: str = (
            f"请输入出哪个手指，可选 {self.available_fingers} 之一，输入 {self.exit_str} 退出游戏: "
        )
        error_prompt: str = "输入错误，请重新输入。"
        while True:
            user_input: str = input(input_prompt)
            if user_input == self.exit_str or user_input in self.available_fingers:
                return user_input
            print(error_prompt)

    def __get_computer_choice(self) -> Finger:
        """
        获取计算机选择的手指。

        :return: 计算机选择的手指
        """
        computer_choice: Finger
        if self.statistics.last_user_win:
            if random.random() < 0.8:
                predicted_fingers: list[Finger] = sorted(
                    self.statistics.user_finger_after_winning,
                    key=self.statistics.user_finger_after_winning.get,
                    reverse=True,
                )[:2]
                predicted_finger = random.choice(seq=predicted_fingers)
                computer_choice = self.__get_winning_finger(
                    predicted_finger=predicted_finger
                )
            else:
                computer_choice = random.choice(seq=self.available_fingers)
        else:
            if random.random() < 0.75:
                remaining_fingers: list[Finger] = [
                    finger
                    for finger in self.available_fingers
                    if finger != self.statistics.last_user_finger
                ]
                predicted_finger: Finger = random.choice(seq=remaining_fingers)
                computer_choice = self.__get_winning_finger(
                    predicted_finger=predicted_finger
                )
            else:
                computer_choice = random.choice(seq=self.available_fingers)
        return computer_choice

    def __get_winning_finger(self, predicted_finger: Finger) -> Finger:
        """
        获取能赢预测手指的手指。

        :param predicted_finger: 预测的手指
        :return: 能赢预测手指的手指
        """
        for win, lose in self.win_lose_pairs:
            if lose == predicted_finger:
                return win
        assert False, "无法找到能赢预测手指的手指, 这部分代码不应该被执行."

    def __judge(self, user_choice: Finger, computer_choice: Finger) -> None:
        """
        判断胜负并更新统计数据。

        :param user_choice: 用户选择的手指
        :param computer_choice: 计算机选择的手指
        """
        win_output: str = "你赢了"
        lose_output: str = "你输了"
        draw_output: str = "平局"

        print(f"计算机选择出 {computer_choice}!")
        if (user_choice, computer_choice) in self.win_lose_pairs:
            print(win_output)
            self.statistics.update_stats(
                user_finger=user_choice, win=True, draw=False, lose=False
            )

        elif (computer_choice, user_choice) in self.win_lose_pairs:
            print(lose_output)
            self.statistics.update_stats(
                user_finger=user_choice, win=False, draw=False, lose=True
            )

        else:
            print(draw_output)
            self.statistics.update_stats(
                user_finger=user_choice, win=False, draw=True, lose=False
            )

    class Statistics:
        """
        Statistics 类用于记录游戏的统计数据。
        """

        def __init__(self, available_fingers: list[Finger]) -> None:
            """
            初始化统计数据。

            :param available_fingers: 可用的手指列表
            """
            self.total_turn = 0
            self.win_turn = 0
            self.lose_turn = 0
            self.draw_turn = 0
            self.user_finger_after_winning: dict[Finger, int] = {
                finger: 0 for finger in available_fingers
            }
            self.last_user_finger = None
            self.last_user_win = False

        def update_stats(
            self, user_finger: Finger, win: bool, draw: bool, lose: bool
        ) -> None:
            """
            更新统计数据。

            :param user_finger: 用户出的手指
            :param win: 是否赢了
            :param draw: 是否平局
            :param lose: 是否输了
            """
            self.total_turn += 1
            if draw:
                self.draw_turn += 1
            elif win:
                self.win_turn += 1
            elif lose:
                self.lose_turn += 1
            else:
                assert False, "对局状态出错"
            if self.last_user_win:
                self.user_finger_after_winning[user_finger] += 1
            self.last_user_finger: Finger = user_finger
            self.last_user_win: bool = win

        def to_string(self) -> str:
            """
            获取字符串形式的统计信息

            :return: 字符串形式的游戏统计数据
            """
            return (
                "游戏结束, 统计信息如下:\n"
                + f"\t总游戏数:\t{self.total_turn};"
                + f"\t赢的次数:\t{self.win_turn};"
                + f"\t输的次数:\t{self.lose_turn};"
                + f"\t平局次数:\t{self.draw_turn};"
            )

    def start(self) -> None:
        """
        开始游戏。
        """
        print("压手指游戏开始! 祝你玩的愉快!")
        while True:
            user_input: str = self.__get_input()
            if user_input == self.exit_str:
                return
            self.__judge(
                user_choice=user_input, computer_choice=self.__get_computer_choice()
            )

    def statistic(self) -> Statistics:
        """
        获取游戏统计数据。

        :return: 游戏统计数据
        """
        return self.statistics


if __name__ == "__main__":
    game = FingerGame()
    game.start()
    result: FingerGame.Statistics = game.statistic()
    print(result.to_string())
    waiting_time: float = 5
    print(f"\n程序将在 {waiting_time} 秒后退出.")
    time.sleep(waiting_time)
    sys.exit(0)
