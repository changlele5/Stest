from tkinter import Tk, Label, Entry, Button, StringVar
from numpy import linspace, sqrt, array, ones, sign, all
from pandas import DataFrame
from scipy.stats import t as T_func
from matplotlib import pyplot as plt


class Stest():
    """计算器窗体类"""
    def __init__(self):
        """初始化实例"""
        self.toop = Tk()
        self.display = StringVar()
        self.display.set('Please fill in the blank!')
        self.Demo()

    def input_bar(self, label_text, y_at, x_begin, v):
        bar_width, bar_height, in_gap = 72, 35, 5
        label = Label(self.toop, text=label_text, font=('Times', 11), anchor='e')
        label.place(x=x_begin, y=y_at, width=bar_width, height=bar_height)
        entry = Entry(self.toop, textvariable=v, borderwidth=2, bg="white", font=('Times',16))
        entry.place(x=x_begin+bar_width+in_gap, y=y_at, width=bar_width+20, height=bar_height)

    def Demo(self):
        self.toop.geometry("%dx%d" % (540, 410))
        self.toop.resizable(False, False)
        self.toop.title(u"Test of main effect")

        # 打印框
        label = Label(self.toop, textvariable=self.display, bg='#ffffff', font=('Times', 20))
        label.place(x=0, y=0, width=540, height=80)

        self.vlist = [StringVar() for i in range(15)]
        # 第1行
        i = 0
        text_lines = ['X-mean:', 'X-std:', 'N:']
        self.input_bar(text_lines[0], 100+(i//3)*50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100+(i//3)*50, 180, self.vlist[i+1])
        self.input_bar(text_lines[2], 100+(i//3)*50, 360, self.vlist[i+2])

        # 第2行
        i = 3
        text_lines = ['b1:', 'b2:', 'b3:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第3行
        i = 6
        text_lines = ['s11:', 's22:', 's33:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第4行
        i = 9
        text_lines = ['s12:', 's13:', 's23:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第4行
        i = 12
        text_lines = ['Intercept:', 'Xname:', 'Yname:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 输出按钮
        button = Button(self.toop, text='Run and plot', bg='white', font=('Times', 16), command=self.run)
        button.place(x=220, y=350, width=130, height=45)

        self.toop.mainloop()

    def run(self):
        try:
            result = [float(v.get()) for v in self.vlist[:13]]
        except:
            self.display.set('Input Format Error!')
        xmean, xstd, N = result[0], result[1], result[2]
        b1, b2, b3 = result[3], result[4], result[5]
        s11, s22, s33 = result[6], result[7], result[8]
        s12, s13, s23 = result[9], result[10], result[11]
        intercept_ = result[12]
        Xname, Yname = self.vlist[13].get(), self.vlist[14].get()
        X = linspace(xmean-2*xstd, xmean+2*xstd, 1000)
        b = b1 + 2 * b2 * X + 3 * b3 * (X**2)
        sb = s11 + 4 * (X**2) * s22 + 9 * (X**4) * s33 + 4 * X * s12 + 6 * (X**2) * s13 + 12 * (X**3) * s23
        t_val = b/sqrt(sb)
        results = DataFrame(array([X, b, sb, t_val]).transpose(), columns=['X', 'Simple Slope', 'Variance', 'T-stat'])
        results.to_excel('./results-main.xlsx')

        ### plot
        def fit(X):
            return intercept_ + b1 * X + b2 * (X**2) + b3 * (X**3)

        alpha = 0.05
        T_trim = T_func.ppf(1-alpha/2, N-1)

        b3_sign = sign(b3)

        plt.rc('font', family='Times New Roman')
        fig = plt.figure(figsize=(9, 6), dpi=100)
        ax = fig.add_subplot(111)
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        if 4 * b2 * b2 - 4 * 3 * b3 * b1 <= 0 or b3==0:
            self.display.set('S-relationship is unseen')
            ax.plot(results['X'], fit(results['X']), linestyle='-', linewidth=1, color='black')
        else:
            stationary = [(-2 * b2 - sqrt(4 * b2 * b2 - 4 * 3 * b3 * b1)) / (2 * 3 * b3),
                          (-2 * b2 + sqrt(4 * b2 * b2 - 4 * 3 * b3 * b1)) / (2 * 3 * b3)]
            stage1 = min(stationary)
            stage2 = max(stationary)
            if stage1 < xmean - 2*xstd or stage2 > xmean + 2*xstd:
                self.display.set('S-relationship is insignificant')
                ax.plot(results['X'], fit(results['X']), linestyle='-', linewidth=1, color='black')
            else:
                part1 = results[results['X'] < stage1]
                part2 = results[results['X'] > stage1]
                part2 = part2[part2['X'] < stage2]
                part3 = results[results['X'] > stage2]
                min_ = fit(results['X']).min()
                max_ = fit(results['X']).max()

                ##  stage 1
                threshold_id = 1
                x_low_significant = part1[b3_sign * part1['T-stat'] > T_trim]['X']
                if len(x_low_significant) == 0:
                    ax.plot(part1['X'], fit(part1['X']), linestyle='--', linewidth=1, color='gray')
                else:
                    ax.plot(x_low_significant, fit(x_low_significant), linestyle='-', linewidth=2, color='black')
                    x_low_insignificant = part1[b3_sign * part1['T-stat'] < T_trim]

                    threshold1 = x_low_significant.min()
                    x_low_insignificant1 = x_low_insignificant['X'][x_low_insignificant['X'] < threshold1]
                    if len(x_low_insignificant1) != 0:
                        ax.plot(x_low_insignificant1, fit(x_low_insignificant1), linestyle='--', linewidth=1, color='gray')

                    if threshold1 > xmean - 2*xstd:
                        ax.plot(threshold1 * ones(100), linspace(min_, fit(threshold1), 100), linestyle=':',
                                linewidth=0.5, color='red')
                        ax.text(threshold1, fit(threshold1) + 0.02 * (max_ - min_),
                                'Threshold {}'.format(threshold_id),
                                ha='left', rotation=90, fontsize=9)
                        threshold_id += 1

                    threshold2 = x_low_significant.max()
                    x_low_insignificant2 = x_low_insignificant['X'][x_low_insignificant['X'] > threshold2]
                    if len(x_low_insignificant2) != 0:
                        ax.plot(x_low_insignificant2, fit(x_low_insignificant2), linestyle='--', linewidth=1, color='gray')
                    ax.plot(threshold2 * ones(100), linspace(min_, fit(threshold2), 100), linestyle=':',
                            linewidth=0.5, color='red')
                    ax.text(threshold2, fit(threshold2) + 0.02 * (max_ - min_),
                            'Threshold {}'.format(threshold_id),
                            ha='left', rotation=90, fontsize=9)
                    threshold_id += 1

                ##  stage 2
                x_medium_significant = part2[b3_sign * part2['T-stat'] < -T_trim]['X']
                if len(x_medium_significant) == 0:
                    ax.plot(part2['X'], fit(part2['X']), linestyle='-', linewidth=1, color='gray')
                else:
                    ax.plot(x_medium_significant, fit(x_medium_significant), linestyle='-', linewidth=2, color='black')
                    x_medium_insignificant = part2[b3_sign * part2['T-stat'] > -T_trim]

                    threshold3 = x_medium_significant.min()
                    x_medium_insignificant1 = x_medium_insignificant['X'][x_medium_insignificant['X'] < threshold3]
                    if len(x_medium_insignificant1) != 0:
                        ax.plot(x_medium_insignificant1, fit(x_medium_insignificant1), linestyle='--', linewidth=1,
                                color='gray')
                    ax.plot(threshold3 * ones(100), linspace(min_, fit(threshold3), 100), linestyle=':',
                            linewidth=0.5, color='red')
                    ax.text(threshold3, fit(threshold3) + 0.02 * (max_ - min_),
                            'Threshold {}'.format(threshold_id),
                            ha='left', rotation=90, fontsize=9)
                    threshold_id += 1

                    threshold4 = x_medium_significant.max()
                    x_medium_insignificant2 = x_medium_insignificant['X'][x_medium_insignificant['X'] > threshold4]
                    if len(x_medium_insignificant2) != 0:
                        ax.plot(x_medium_insignificant2, fit(x_medium_insignificant2), linestyle='--', linewidth=1,
                                color='gray')
                    ax.plot(threshold4 * ones(100), linspace(min_, fit(threshold4), 100), linestyle=':',
                            linewidth=0.5, color='red')
                    ax.text(threshold4, fit(threshold4) + 0.02 * (max_ - min_),
                            'Threshold {}'.format(threshold_id),
                            ha='left', rotation=90, fontsize=9)
                    threshold_id += 1

                ##  stage 3
                x_high_significant  = part3[b3_sign * part3['T-stat'] > T_trim]['X']

                if len(x_high_significant) == 0:
                    ax.plot(part3['X'], fit(part3['X']), linestyle='--', linewidth=1, color='gray')
                else:
                    ax.plot(x_high_significant, fit(x_high_significant), linestyle='-', linewidth=2, color='black')
                    x_high_insignificant = part3[b3_sign * part3['T-stat'] < T_trim]

                    threshold5 = x_high_significant.min()
                    x_high_insignificant1 = x_high_insignificant['X'][x_high_insignificant['X'] < threshold5]
                    if len(x_high_insignificant1) != 0:
                        ax.plot(x_high_insignificant1, fit(x_high_insignificant1), linestyle='--', linewidth=1, color='gray')
                    ax.plot(threshold5 * ones(100), linspace(min_, fit(threshold5), 100), linestyle=':',
                            linewidth=0.5, color='red')
                    ax.text(threshold5, fit(threshold5) + 0.02 * (max_ - min_),
                            'Threshold {}'.format(threshold_id),
                            ha='left', rotation=90, fontsize=9)
                    threshold_id += 1

                    threshold6 = x_high_significant.max()
                    x_high_insignificant2 = x_high_insignificant['X'][x_high_insignificant['X'] > threshold6]
                    if len(x_high_insignificant2) != 0:
                        ax.plot(x_high_insignificant2, fit(x_high_insignificant2), linestyle='--', linewidth=1,
                                color='gray')
                    if threshold6 < xmean + 2 * xstd:
                        ax.plot(threshold6 * ones(100), linspace(min_, fit(threshold6), 100), linestyle=':',
                                linewidth=0.5, color='red')
                        ax.text(threshold6, fit(threshold6) + 0.02 * (max_ - min_),
                                'Threshold {}'.format(threshold_id),
                                ha='left', rotation=90, fontsize=9)
                        threshold_id += 1

                ax.plot(linspace(xmean-2*xstd, xmean+2*xstd, 100), min_*ones(100),
                        linestyle=':', linewidth=0.5, color='gray', alpha=0.5)

        ax.grid(True, axis='y', linestyle='-', color='gray', alpha=0.2)
        ax.set_xticks([xmean - 2 * xstd, xmean - xstd, xmean, xmean + xstd, xmean + 2 * xstd])
        ax.set_xticklabels(['-2SD', '-SD', '0', 'SD', '2SD'], rotation=0, fontsize=12)
        plt.xlabel(Xname, fontsize=13)
        plt.ylabel(Yname, fontsize=13)
        plt.savefig('S-fit.png', dpi=300)
        plt.show()

        self.display.set('Output to results-main.xlsx!\nGraph at S-fit.png.')

class Stest_with_moderator():
    """计算器窗体类"""
    def __init__(self):
        """初始化实例"""
        self.toop = Tk()
        self.display = StringVar()
        self.display.set('Please fill in the blank!')
        self.Demo()

    def input_bar(self, label_text, y_at, x_begin, v):
        bar_width, bar_height, in_gap = 72, 35, 5
        label = Label(self.toop, text=label_text, font=('Times', 11), anchor='e')
        label.place(x=x_begin, y=y_at, width=bar_width, height=bar_height)
        entry = Entry(self.toop, textvariable=v, borderwidth=2, bg="white", font=('Times',16))
        entry.place(x=x_begin+bar_width+in_gap, y=y_at, width=bar_width+20, height=bar_height)

    def Demo(self):
        self.toop.geometry("%dx%d" % (540, 710))
        self.toop.resizable(False, False)
        self.toop.title(u"Test of moderating effect")

        # 打印框
        label = Label(self.toop, textvariable=self.display, bg='#ffffff', font=('Times', 20))
        label.place(x=0, y=0, width=540, height=80)

        self.vlist = [StringVar() for i in range(35)]
        # 第1行
        i = 0
        text_lines = ['X-mean:', 'X-std:', 'Z-std:']
        self.input_bar(text_lines[0], 100+(i//3)*50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100+(i//3)*50, 180, self.vlist[i+1])
        self.input_bar(text_lines[2], 100+(i//3)*50, 360, self.vlist[i+2])

        # 第2行
        i = 3
        text_lines = ['b1:', 'b2:', 'b3:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第3行
        i = 6
        text_lines = ['b5:', 'b6:', 'b7:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第4行
        i = 9
        text_lines = ['s11:', 's22:', 's33:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第5行
        i = 12
        text_lines = ['s55:', 's66:', 's77:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第6行
        i = 15
        text_lines = ['s12:', 's13:', 's15:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第7行
        i = 18
        text_lines = ['s16:', 's17:', 's23:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第8行
        i = 21
        text_lines = ['s25:', 's26:', 's27:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第9行
        i = 24
        text_lines = ['s35:', 's36:', 's37:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第10行
        i = 27
        text_lines = ['s56:', 's57:', 's67:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第11行
        i = 30
        text_lines = ['N:', 'Intercept:', 'Xname:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 180, self.vlist[i + 1])
        self.input_bar(text_lines[2], 100 + (i // 3) * 50, 360, self.vlist[i + 2])

        # 第12行
        i = 33
        text_lines = ['Moderator:', 'Yname:']
        self.input_bar(text_lines[0], 100 + (i // 3) * 50, 0, self.vlist[i])
        self.input_bar(text_lines[1], 100 + (i // 3) * 50, 360, self.vlist[i+1])

        # 输出按钮
        button = Button(self.toop, text='Run and plot', bg='white', font=('Times', 16), command=self.run)
        button.place(x=220, y=650, width=130, height=45)

        self.toop.mainloop()

    def run(self):
        try:
            result = [float(v.get()) for v in self.vlist[:32]]
        except:
            self.display.set('Format Error of Input!')
        xmean, xstd, Zstd = result[0], result[1], result[2]
        b1, b2, b3 = result[3], result[4], result[5]
        b5, b6, b7 = result[6], result[7], result[8]
        s11, s22, s33 = result[9], result[10], result[11]
        s55, s66, s77 = result[12], result[13], result[14]
        s12, s13, s15 = result[15], result[16], result[17]
        s16, s17, s23 = result[18], result[19], result[20]
        s25, s26, s27 = result[21], result[22], result[23]
        s35, s36, s37 = result[24], result[25], result[26]
        s56, s57, s67 = result[27], result[28], result[29]
        N = result[30]
        intercept_ = result[31]
        Xname, Moderator, Yname = self.vlist[32].get(), self.vlist[33].get(), self.vlist[34].get()

        def fit(X):
            return b1*X + b5*Z*X + b2*(X**2) + b6*Z*(X**2) + b3*(X**3) + b7*Z*(X**3) + intercept_

        alpha = 0.05
        T_trim = T_func.ppf(1 - alpha / 2, N - 1)

        plt.rc('font', family='Times New Roman')
        fig = plt.figure(figsize=(9, 6), dpi=100)
        ax = plt.axes([0.11, 0.13, 0.8, 0.75])
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)

        Z = Zstd
        def run_and_plot(Z, ax, mode_name='High Moderator', color='red', text_mode='top', linestyle='-'):
            ### run
            X = linspace(xmean-2*xstd, xmean+2*xstd, 1000)
            b = b1 + b5*Z + 2*b2*X + 2*b6*Z*X + 3*b3*(X**2) + 3*b7*Z*(X**2)
            sb = s11 + 4*(X**2)*s22 + 9*(X**4)*s33 + s55*(Z**2) + 4*s66*(X**2)*(Z**2) + 4*s12*X + 6*s13*(X**2) + 2*s15*Z
            + 4*s16*X*Z + 6*s17*(X**2)*Z + 12*s23*(X**3) + 4*s25*X*Z + 8*s26*(X**2)*Z + 12*s27*(X**3)*Z + 6*s35*(X**2)*Z
            + 12*s36*(X**3)*Z + 18*s37*(X**4)*Z + 4*s56*X*(Z**2) + 6*s57*(X**2)*(Z**2) + 12*s67*(X**3)*(Z**2)
            t = b/sqrt(sb)
            results = DataFrame(array([X, b, sb, t]).transpose(), columns=['X', 'Simple Slope', 'Variance', 'T-stat'])
            results.to_excel('./results-moderating-' + mode_name +'.xlsx')

            ### plot
            b1_ = b1 + b5*Z
            b2_ = 2*b2 + 2*b6*Z
            b3_ = 3*b3 + 3*b7*Z
            b3_sign = sign(b3_)

            if 4 * b2_ * b2_ - 4 * 3 * b3_ * b1_ <= 0 or b3_==0:
                self.display.set('S-relationship is unseen when Z is '+mode_name)
                ax.plot(results['X'], fit(results['X']), linestyle=linestyle, linewidth=1, color=color, label=mode_name)
            else:
                stationary = [(-2 * b2_ - sqrt(4 * b2_ * b2_ - 4 * 3 * b3_ * b1_)) / (2 * 3 * b3_),
                              (-2 * b2_ + sqrt(4 * b2_ * b2_ - 4 * 3 * b3_ * b1_)) / (2 * 3 * b3_)]
                stage1 = min(stationary)
                stage2 = max(stationary)
                if stage1 < xmean - 2*xstd or stage2 > xmean + 2*xstd:
                    self.display.set('S-relationship is insignificant when Z is '+mode_name, label=mode_name)
                    ax.plot(results['X'], fit(results['X']), linestyle=linestyle, linewidth=1, color=color)
                else:
                    label_encode = 0
                    stationary = [(-2 * b2_ - sqrt(4 * b2_ * b2_ - 4 * 3 * b3_ * b1_)) / (2 * 3 * b3_),
                                  (-2 * b2_ + sqrt(4 * b2_ * b2_ - 4 * 3 * b3_ * b1_)) / (2 * 3 * b3_)]
                    stage1 = min(stationary)
                    stage2 = max(stationary)

                    part1 = results[results['X'] < stage1]
                    part2 = results[results['X'] > stage1]
                    part2 = part2[part2['X'] < stage2]
                    part3 = results[results['X'] > stage2]
                    min_ = fit(results['X']).min()
                    max_ = fit(results['X']).max()

                    if text_mode == 'top':
                        yl = max_
                        yp = max_ + 0.01 * (max_ - min_)
                        rotation = 90
                        ha = 'left'
                        va = 'bottom'
                    elif text_mode == 'bottom':
                        yl = min_
                        yp = min_ - 0.01 * (max_ - min_)
                        rotation = 270
                        ha = 'right'
                        va = 'top'
                    else:
                        yl = 0
                        yp = 0 + 0.01 * (max_ - min_)
                        rotation = 90
                        ha = 'left'
                        va = 'bottom'
                    ##  stage 1
                    threshold_id = 1
                    x_low_significant = part1[b3_sign * part1['T-stat'] > T_trim]['X']
                    if len(x_low_significant) == 0:
                        ax.plot(part1['X'], fit(part1['X']), linestyle=linestyle, linewidth=1,
                                color=color, alpha=0.3)
                    else:
                        if label_encode == 0:
                            ax.plot(x_low_significant, fit(x_low_significant), linestyle=linestyle,
                                    linewidth=2, color=color, label=mode_name)
                            label_encode = 1
                        else:
                            ax.plot(x_low_significant, fit(x_low_significant), linestyle=linestyle,
                                    linewidth=2, color=color)
                        x_low_insignificant = part1[b3_sign * part1['T-stat'] < T_trim]

                        threshold1 = x_low_significant.min()
                        x_low_insignificant1 = x_low_insignificant['X'][x_low_insignificant['X'] < threshold1]
                        if len(x_low_insignificant1) != 0:
                            ax.plot(x_low_insignificant1, fit(x_low_insignificant1), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        if threshold1 > xmean - 2 * xstd:
                            ax.plot(threshold1 * ones(100), linspace(yl, fit(threshold1), 100), linestyle=':',
                                    linewidth=0.5, color=color)
                            ax.text(threshold1, yp,
                                    'Threshold {}'.format(threshold_id), color=color, va=va,
                                    ha=ha, rotation=rotation, fontsize=9)
                            threshold_id += 1

                        threshold2 = x_low_significant.max()
                        x_low_insignificant2 = x_low_insignificant['X'][x_low_insignificant['X'] > threshold2]
                        if len(x_low_insignificant2) != 0:
                            ax.plot(x_low_insignificant2, fit(x_low_insignificant2), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        ax.plot(threshold2 * ones(100), linspace(yl, fit(threshold2), 100), linestyle=':',
                                linewidth=0.5, color=color)
                        ax.text(threshold2, yp,
                                'Threshold {}'.format(threshold_id), color=color, va=va,
                                ha=ha, rotation=rotation, fontsize=9)
                        threshold_id += 1

                    ##  stage 2
                    x_medium_significant = part2[b3_sign * part2['T-stat'] < -T_trim]['X']
                    if len(x_medium_significant) == 0:
                        ax.plot(part2['X'], fit(part2['X']), linestyle=linestyle, linewidth=1, color=color, alpha=0.3)
                    else:
                        if label_encode == 0:
                            ax.plot(x_medium_significant, fit(x_medium_significant), linestyle=linestyle, linewidth=2,
                                    color=color, label=mode_name)
                            label_encode = 1
                        else:
                            ax.plot(x_medium_significant, fit(x_medium_significant), linestyle=linestyle, linewidth=2,
                                    color=color)

                        x_medium_insignificant = part2[b3_sign * part2['T-stat'] > -T_trim]

                        threshold3 = x_medium_significant.min()
                        x_medium_insignificant1 = x_medium_insignificant['X'][x_medium_insignificant['X'] < threshold3]
                        if len(x_medium_insignificant1) != 0:
                            ax.plot(x_medium_insignificant1, fit(x_medium_insignificant1), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        ax.plot(threshold3 * ones(100), linspace(yl, fit(threshold3), 100), linestyle=':',
                                linewidth=0.5, color=color)
                        ax.text(threshold3, yp,
                                'Threshold {}'.format(threshold_id), color=color, va=va,
                                ha=ha, rotation=rotation, fontsize=9)
                        threshold_id += 1

                        threshold4 = x_medium_significant.max()
                        x_medium_insignificant2 = x_medium_insignificant['X'][x_medium_insignificant['X'] > threshold4]
                        if len(x_medium_insignificant2) != 0:
                            ax.plot(x_medium_insignificant2, fit(x_medium_insignificant2), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        ax.plot(threshold4 * ones(100), linspace(yl, fit(threshold4), 100), linestyle=':',
                                linewidth=0.5, color=color)
                        ax.text(threshold4, yp,
                                'Threshold {}'.format(threshold_id), color=color, va=va,
                                ha=ha, rotation=rotation, fontsize=9)
                        threshold_id += 1

                    ##  stage 3
                    x_high_significant  = part3[b3_sign * part3['T-stat'] > T_trim]['X']

                    if len(x_high_significant) == 0:
                        if label_encode == 0:
                            ax.plot(part3['X'], fit(part3['X']), linestyle=linestyle, linewidth=1,
                                    color=color, alpha=0.3, label=mode_name)
                            label_encode = 1
                        else:
                            ax.plot(part3['X'], fit(part3['X']), linestyle=linestyle, linewidth=1,
                                    color=color, alpha=0.3)

                    else:
                        if label_encode == 0:
                            ax.plot(x_high_significant, fit(x_high_significant), linestyle=linestyle, linewidth=2,
                                    color=color, label=mode_name)
                            label_encode = 1
                        else:
                            ax.plot(x_high_significant, fit(x_high_significant), linestyle=linestyle, linewidth=2,
                                    color=color)

                        x_high_insignificant = part3[b3_sign * part3['T-stat'] < T_trim]

                        threshold5 = x_high_significant.min()
                        x_high_insignificant1 = x_high_insignificant['X'][x_high_insignificant['X'] < threshold5]
                        if len(x_high_insignificant1) != 0:
                            ax.plot(x_high_insignificant1, fit(x_high_insignificant1), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        ax.plot(threshold5 * ones(100), linspace(yl, fit(threshold5), 100), linestyle=':',
                                linewidth=0.5, color=color)
                        ax.text(threshold5, yp,
                                'Threshold {}'.format(threshold_id), color=color, va=va,
                                ha=ha, rotation=rotation, fontsize=9)
                        threshold_id += 1

                        threshold6 = x_high_significant.max()
                        x_high_insignificant2 = x_high_insignificant['X'][x_high_insignificant['X'] > threshold6]
                        if len(x_high_insignificant2) != 0:
                            ax.plot(x_high_insignificant2, fit(x_high_insignificant2), linestyle=linestyle,
                                    linewidth=1, color=color, alpha=0.3)
                        if threshold6 < xmean + 2 * xstd:
                            ax.plot(threshold6 * ones(100), linspace(yl, fit(threshold6), 100), linestyle=':',
                                    linewidth=0.5, color=color)
                            ax.text(threshold6, yp,
                                    'Threshold {}'.format(threshold_id), color=color, va=va,
                                    ha=ha, rotation=rotation, fontsize=9)
                            threshold_id += 1
                    if text_mode == 'top':
                        ax.plot(linspace(xmean - 2 * xstd, xmean + 2 * xstd, 100), max_ * ones(100),
                                linestyle='-.', linewidth=0.6, color=color)
                    elif text_mode == 'bottom':
                        ax.plot(linspace(xmean - 2 * xstd, xmean + 2 * xstd, 100), min_ * ones(100),
                                linestyle='-.', linewidth=0.6, color=color)
                    else:
                        ax.plot(linspace(xmean - 2 * xstd, xmean + 2 * xstd, 100), 0. * ones(100),
                                linestyle='-.', linewidth=0.6, color=color)

            return ax

        Z = Zstd
        ax = run_and_plot(Z, ax, mode_name='High '+Moderator, color='red', text_mode='top', linestyle='-')
        Z = -Zstd
        ax = run_and_plot(Z, ax, mode_name='Low '+Moderator, color='blue', text_mode='bottom', linestyle='--')

        ax.grid(True, axis='y', linestyle='-', color='gray', alpha=0.2)
        ax.set_xticks([xmean - 2 * xstd, xmean - xstd, xmean, xmean + xstd, xmean + 2 * xstd])
        ax.set_xticklabels(['-2SD', '-SD', '0', 'SD', '2SD'], rotation=0, fontsize=12)
        plt.xlabel(Xname, fontsize=13)
        plt.ylabel(Yname, fontsize=13)
        ax.legend(bbox_to_anchor=(1.01,-0.173),loc='lower right')
        plt.savefig('S-moderating.png', dpi=300)
        plt.show()

        self.display.set('Output to results-moderating.xlsx!\nGraph at S-moderating.png.')


if __name__ == "__main__":
    root = Tk()
    root.geometry("%dx%d" % (540, 400))
    root.resizable(False, False)
    root.title(u"The testing for S-Shaped Relationship")

    def main_test():
        root.destroy()
        Stest()
    button1 = Button(root, text='Main effect test', bg='white', font=('Times', 16), command=main_test)
    button1.place(x=120, y=100, width=300, height=80)

    def moderating_test():
        root.destroy()
        Stest_with_moderator()
    button2 = Button(root, text='Moderating effect test', bg='white', font=('Times', 16), command=moderating_test)
    button2.place(x=120, y=220, width=300, height=80)

    root.mainloop()
