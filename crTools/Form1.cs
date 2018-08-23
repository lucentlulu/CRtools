using CRAPI;
using crTools.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.CellTypes;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using System.Configuration;
using System.IO;

namespace crTools
{
    public enum GridHeader
    {
        [StringValue("序号")]
        id = 0,
        [StringValue("玩家")]
        name = 1,
        [StringValue("职位")]
        role = 2,
        [StringValue("杯数")]
        score = 3,
        [StringValue("国王")]
        level = 4,
        [StringValue("竞技场")]
        arena = 5,
        [StringValue("每周捐出")]
        outcards = 6,
        [StringValue("每周接受")]
        incards = 7,
        [StringValue("每周结算")]
        deltacards = 8,
        [StringValue("捐赠总量")]
        totaloutcards = 9,
        [StringValue("集卡总量")]
        clanCards = 10,
        [StringValue("卡牌缺")]
        cards = 11,
        [StringValue("近期常用")]
        favoriteCard = 12,
        [StringValue("三冠胜场")]
        wins3 = 13,
        [StringValue("胜场")]
        wins = 14,
        [StringValue("负场")]
        losses = 15,
        [StringValue("平局")]
        draws = 16,
        [StringValue("部落胜场")]
        clanwins = 17,
        [StringValue("胜率")]
        winsPercent = 18,
        [StringValue("部落集卡")]
        cardearns = 19,
        [StringValue("部落场数")]
        cardbattles = 20,
        [StringValue("部落胜场")]
        cardwins = 21,

        [StringValue("1场数")]
        cardbattles1 = 24,
        [StringValue("1胜场")]
        cardwins1 = 23,
        [StringValue("1卡牌")]
        cardearns1 = 22,
        [StringValue("2场数")]
        cardbattles2 = 27,
        [StringValue("2胜场")]
        cardwins2 = 26,
        [StringValue("2卡牌")]
        cardearns2 = 25,
        [StringValue("3场数")]
        cardbattles3 = 30,
        [StringValue("3胜场")]
        cardwins3 = 29,
        [StringValue("3卡牌")]
        cardearns3 = 28,
        [StringValue("4场数")]
        cardbattles4 = 33,
        [StringValue("4胜场")]
        cardwins4 = 32,
        [StringValue("4卡牌")]
        cardearns4 = 31,
        [StringValue("5场数")]
        cardbattles5 = 36,
        [StringValue("5胜场")]
        cardwins5 = 35,
        [StringValue("5卡牌")]
        cardearns5 = 34,
        [StringValue("6场数")]
        cardbattles6 = 39,
        [StringValue("6胜场")]
        cardwins6 = 38,
        [StringValue("6卡牌")]
        cardearns6 = 37,
        [StringValue("7场数")]
        cardbattles7 = 42,
        [StringValue("7胜场")]
        cardwins7 = 41,
        [StringValue("7卡牌")]
        cardearns7 = 40,
        [StringValue("8场数")]
        cardbattles8 = 45,
        [StringValue("8胜场")]
        cardwins8 = 44,
        [StringValue("8卡牌")]
        cardearns8 = 43,
        [StringValue("9场数")]
        cardbattles9 = 48,
        [StringValue("9胜场")]
        cardwins9 = 47,
        [StringValue("9卡牌")]
        cardearns9 = 46,
        [StringValue("10场数")]
        cardbattles10 = 51,
        [StringValue("10胜场")]
        cardwins10 = 50,
        [StringValue("10卡牌")]
        cardearns10 = 49,
        [StringValue("胜场小计")]
        clanwin = 52,
        [StringValue("胜率小计")]
        clanwinpercent = 53,

        [StringValue("下1个")]
        chest1 = 54,
        [StringValue("下2个")]
        chest2 = 55,
        [StringValue("下3个")]
        chest3 = 56,
        [StringValue("下4个")]
        chest4 = 57,
        [StringValue("下5个")]
        chest5 = 58,
        [StringValue("下6个")]
        chest6 = 59,
        [StringValue("下7个")]
        chest7 = 60,
        [StringValue("下8个")]
        chest8 = 61,
        [StringValue("下9个")]
        chest9 = 62,
        [StringValue("巨型宝箱")]
        giant = 63,
        [StringValue("神奇宝箱")]
        magical = 64,
        [StringValue("超级神奇宝箱")]
        supermagical = 65,
        [StringValue("史诗宝箱")]
        epic = 66,
        [StringValue("传奇宝箱")]
        legendary = 67,
        [StringValue("")]
        totalcolumns = 68,
    };

    public partial class Form1 : Form
    {
        private Wrapper clashRoyale = null;
        private ClanWarInfo clanwar = null;
        private ClanWarsRecord[] clanwars = null;
        private Player[] players = null;
        //private string filename = "cr.rgf";
        //private static int TagLen = 8;
        private static int MaxChest = 9;
        private static int MaxMember = 50;
        private static string TagName = "ClanTag";
        private static int SumRow = MaxMember + 2;
        private static ushort[] HeaderWidth = new ushort[(int)GridHeader.totalcolumns]
        {
            35,150, 50, 40, 40,120, 65, 65, 65, 65,
            65, 50, 65, 65, 50, 50, 50, 65, 55, 65, 65, 65, 65,
            60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60,
            60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 65, 65,
            70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 70, 100, 70, 70
        };
        private static string[] chests = new string[]
        {   "普通宝箱","白银宝箱","黄金宝箱","巨型宝箱","神奇宝箱","超级神奇宝箱","传奇宝箱","史诗宝箱","皇冠宝箱"};
        private static string[] arenas = new string[]
        {
            "训练营","哥布林竞技场","骷髅深渊","野蛮人擂台","皮卡超人乐园","法术幽谷","建筑工人工坊","皇家竞技场","冰封之巅","丛林竞技场","野猪山脉","电磁峡谷","传奇竞技场",
            "挑战者联赛 I","挑战者联赛 II","挑战者联赛 III","大师联赛 I","大师联赛 II","大师联赛 III","冠军联赛","超级冠军联赛","终极冠军联赛"
        };
        private static string[,] cards = new string[87, 3]
        {
            {"archers",             "Archers",              "弓箭手"},
            {"arrows",              "Arrows",               "箭雨"},
            {"baby-dragon",         "Baby Dragon",          "龙宝"},
            {"balloon",             "Balloon",              "气球"},
            {"bandit",              "Bandit",               "刺客"},
            {"barbarian-barrel",    "Barbarian Barrel",     "滚筒兵"},
            {"barbarian-hut",       "Barbarian Hut",        "小木屋"},
            {"barbarians",          "barbarians",           "黄毛"},
            {"bats",                "Bats",                 "蝙蝠"},
            {"battle-ram",          "Battle Ram",           "攻城槌"},
            {"bomb-tower",          "Bomb Tower",           "投弹塔"},
            {"bomber",              "Bomber",               "投弹兵"},
            {"bowler",              "Bowler",               "蓝胖"},
            {"cannon-cart",         "Cannon Cart",          "加农炮车"},
            {"cannon",              "Cannon",               "加农炮塔"},
            {"clone",               "Clone",                "克隆"},
            {"dark-prince",         "Dark Prince",          "黑暗王子"},
            {"dart-goblin",         "Dart Goblin",          "吹箭兵"},
            {"electro-wizard",      "Electro Wizard",       "电法"},
            {"elite-barbarians",    "Elite Barbarians",     "精锐兵"},
            {"elixir-collector",    "Elixir Collector",     "采集器"},
            {"executioner",         "Executioner",          "飞斧兵"},
            {"fire-spirits",        "Fire Spirits",         "火精灵"},
            {"fireball",            "Fire Ball",            "火球"},
            {"flying-machine",      "Flying Machine",       "无人机"},
            {"freeze",              "Freeze",               "冰冻"},
            {"furnace",             "Furnace",              "熔炉"},
            {"giant-skeleton",      "Giant Skeleton",       "骷髅巨人"},
            {"giant-snowball",      "Giant SnowBall",       "雪球"},
            {"giant",               "Giant",                "胖子"},
            {"goblin-barrel",       "Goblin Barrel",        "全家桶"},
            {"goblin-gang",         "Goblin Gang",          "哥布林团伙"},
            {"goblin-hut",          "Goblin Hut",           "茅草屋"},
            {"goblins",             "Goblins",              "哥布林"},
            {"golem",               "Golem",                "石头人"},
            {"graveyard",           "Graveyard",            "墓地"},
            {"guards",              "Guards",               "士兵"},
            {"heal",                "Heal",                 "治愈"},
            {"hog-rider",           "Hog Rider",            "野猪骑士"},
            {"hunter",              "Hunter",               "猎人"},
            {"ice-golem",           "Ice Golem",            "冰胖"},
            {"ice-spirit",          "Ice Spirit",           "冰精灵"},
            {"ice-wizard",          "Ice Wizard",           "冰法"},
            {"inferno-dragon",      "Inferno Dragon",       "地狱龙"},
            {"inferno-tower",       "Inferno Tower",        "地狱塔"},
            {"knight",              "Knight",               "骑士"},
            {"lava-hound",          "Lava Hound",           "天狗"},
            {"lightning",           "Lightning",            "大闪"},
            {"lumberjack",          "Lumberjack",           "樵夫"},
            {"magic-archer",        "Magic Archer",         "神箭手"},
            {"mega-knight",         "Mega Knight",          "超骑"},
            {"mega-minion",         "Mega Minion",          "铁苍蝇"},
            {"miner",               "Miner",                "矿工"},
            {"mini-pekka",          "Mini P.E.K.K.A",       "小皮卡"},
            {"minion-horde",        "Minion Horde",         "苍蝇群"},
            {"minions",             "Minions",              "小苍蝇"},
            {"mirror",              "Mirror",               "镜像"},
            {"mortar",              "Mortar",               "迫击炮"},
            {"musketeer",           "Musketeer",            "枪兵"},
            {"night-witch",         "Night Witch",          "暗夜女巫"},
            {"pekka",               "P.E.K.K.A",            "大皮卡"},
            {"poison",              "Poison",               "毒药"},
            {"prince",              "Prince",               "王子"},
            {"princess",            "Princess",             "公主"},
            {"rage",                "Rage",                 "狂暴"},
            {"rascals",             "Rascals",              "绿林团伙"},
            {"rocket",              "Rocket",               "火箭"},
            {"royal-ghost",         "Royal Ghost",          "皇家幽灵"},
            {"royal-giant",         "Royal Giant",          "皇家巨人"},
            {"royal-hogs",          "Royal Hogs",           "皇家野猪"},
            {"royal-recruits",      "Royal Recruits",       "皇家护卫"},
            {"skeleton-army",       "Skeleton Army",        "骷髅群"},
            {"skeleton-barrel",     "Skeleton Barrel",      "骷髅气球"},
            {"Skeletons",           "Skeletons",            "小骷髅"},
            {"sparky",              "Sparky",               "电磁炮"},
            {"spear-goblins",       "Spear Goblins",        "投矛兵"},
            {"tesla",               "Tesla",                "电塔"},
            {"the-log",             "The Log",              "滚木"},
            {"three-musketeers",    "Three Musketeers",     "三枪"},
            {"tombstone",           "Tombstone",            "墓碑"},
            {"tornado",             "Tornado",              "飓风"},
            {"valkyrie",            "Valkyrie",             "瓦基里"},
            {"witch",               "Witch",                "女巫"},
            {"wizard",              "Wizard",               "火法"},
            {"x-bow",               "X-Bow",                "连弩"},
            {"zap",                 "Zap",                  "小闪电"},
            {"zappies",             "Zappies",              "电击车"}
        };
        public Form1(ref Wrapper clashroyale)
        {
            InitializeComponent();

            clashRoyale = clashroyale;
            labelDate.Text = "查询时间：" + DateTime.Now.ToLongDateString();
        }

        private void FillLocation()
        {
            comboBoxRegion.Items.Clear();
            for (Wrapper.Locations i = Wrapper.Locations._EU; i <= Wrapper.Locations.ZW; i++)
            {
                comboBoxRegion.Items.Add(i.ToString());
            }
            comboBoxRegion.SelectedItem = "CN";
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            buttonQuery.Enabled = (clashRoyale != null);
            //buttonSave.Enabled = checkBoxDatabase.Checked;

            string strTag = ConfigurationManager.AppSettings[TagName];
            if (!string.IsNullOrEmpty(strTag)) comboBoxClans.Text = strTag;
            FillLocation();
            SetGridHeader();

        }
        private void SetGridHeader()
        {
            reoGridControl.Reset();
            reoGridControl.CurrentWorksheet.SetRows(MaxMember + 8);
            reoGridControl.CurrentWorksheet.SetCols((int)GridHeader.totalcolumns);
            reoGridControl.CurrentWorksheet.SetRowsHeight(0, 1, 40);

            for (GridHeader gh = GridHeader.id; gh < GridHeader.totalcolumns; gh++)
            {
                if (checkBoxRecord.Checked && gh >= GridHeader.cardbattles1)
                    reoGridControl.CurrentWorksheet[0, (int)gh] = StringEnum.GetStringValue(gh);
                else if (checkBoxWar.Checked && gh >= GridHeader.cardbattles)
                    reoGridControl.CurrentWorksheet[0, (int)gh] = StringEnum.GetStringValue(gh);
                else
                    reoGridControl.CurrentWorksheet[0, (int)gh] = StringEnum.GetStringValue(gh);
                reoGridControl.CurrentWorksheet.SetColumnsWidth((int)gh, 1, HeaderWidth[(int)gh]);
            }
            if (checkBoxViewCardImage.Checked)
            {
                //reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.favoriteCard, 1, (ushort)imageListCards.ImageSize.Width);
                reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.chest1, MaxChest, (ushort)imageListChests.ImageSize.Width);  // * imageListCards.ImageSize.Height / imageListChests.ImageSize.Height));
                reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.favoriteCard, 1, (ushort)(imageListCards.ImageSize.Width * imageListChests.ImageSize.Height / imageListCards.ImageSize.Height));
                reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.arena, 1, (ushort)(imageListArena.ImageSize.Width * imageListChests.ImageSize.Height / imageListArena.ImageSize.Height));
            }
            // 首行
            reoGridControl.DoAction(new SetRangeStyleAction(0, 0, 1, reoGridControl.CurrentWorksheet.ColumnCount,
                                    new WorksheetRangeStyle
                                    {
                                        Flag = PlainStyleFlag.TextWrap | PlainStyleFlag.AlignAll | PlainStyleFlag.BackColor | PlainStyleFlag.FontStyleBold | PlainStyleFlag.TextColor,
                                        Bold = true,
                                        TextColor = Color.Blue,
                                        BackColor = Color.DarkGray,
                                        HAlign = ReoGridHorAlign.Center,
                                        VAlign = ReoGridVerAlign.Middle,
                                        TextWrapMode = TextWrapMode.WordBreak,
                                    }));
            // 首行
            /*
            reoGridControl.DoAction(new SetRangeBorderAction(new RangePosition(0, 0, reoGridControl.CurrentWorksheet.RowCount, reoGridControl.CurrentWorksheet.ColumnCount),
                                        BorderPositions.Outside,
                                    new RangeBorderStyle
                                    {
                                        Color = Color.Black,
                                        Style = BorderLineStyle.BoldSolid,
                                    }));
            */
            // 锁定 标题栏， 序号和姓名
            reoGridControl.CurrentWorksheet.FreezeToCell(1, 2);
            // 捐牌序列
            reoGridControl.CurrentWorksheet.AddOutline(RowOrColumn.Column, (int)GridHeader.outcards, GridHeader.favoriteCard - GridHeader.outcards);
            // 当前战斗
            reoGridControl.CurrentWorksheet.AddOutline(RowOrColumn.Column, (int)GridHeader.cardearns, GridHeader.cardwins - GridHeader.cardearns);
            // 战斗历史
            reoGridControl.CurrentWorksheet.AddOutline(RowOrColumn.Column, (int)GridHeader.cardearns1, GridHeader.clanwinpercent - GridHeader.cardearns1);
            // 宝箱序列
            reoGridControl.CurrentWorksheet.AddOutline(RowOrColumn.Column, (int)GridHeader.chest1, GridHeader.legendary - GridHeader.chest1);
            // 设置胜率的显示格式
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.name] = "合计";
            //reoGridControl.CurrentWorksheet.SetRangeDataFormat(new RangePosition(SumRow, (int)GridHeader.level, 1, GridHeader.totalcolumns - GridHeader.level), unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number, null);
            reoGridControl.CurrentWorksheet.SetRangeDataFormat(new RangePosition(1, (int)GridHeader.winsPercent, reoGridControl.CurrentWorksheet.RowCount, 1), unvell.ReoGrid.DataFormat.CellDataFormatFlag.Percent, null);
            reoGridControl.CurrentWorksheet.SetRangeDataFormat(new RangePosition(1, (int)GridHeader.clanwinpercent, reoGridControl.CurrentWorksheet.RowCount, 1), unvell.ReoGrid.DataFormat.CellDataFormatFlag.Percent, null);
            // 设置排序
            /*
            reoGridControl.CurrentWorksheet.CreateColumnFilter(new RangePosition(1, (int)GridHeader.winsPercent, MaxMember, 1));
            reoGridControl.CurrentWorksheet.CreateColumnFilter(new RangePosition(1, (int)GridHeader.cardwins, MaxMember, 1));
            reoGridControl.CurrentWorksheet.CreateColumnFilter(new RangePosition(1, (int)GridHeader.role, MaxMember, 1));
            reoGridControl.CurrentWorksheet.CreateColumnFilter(new RangePosition(1, (int)GridHeader.winsPercent, MaxMember, 1));
            */
            // 隐藏 row: A - Z  col: 1 - n   
            reoGridControl.CurrentWorksheet.SetSettings(WorksheetSettings.View_ShowFrozenLine | WorksheetSettings.View_ShowColumnHeader | WorksheetSettings.View_ShowRowHeader, false);
        }

        private void setCheckBoxState(bool value)
        {
            progressBarQuery.Visible = !value;
            checkBoxRecord.Enabled = value;
            checkBoxChest.Enabled = value;
            checkBoxWar.Enabled = value;
            checkBoxViewCardImage.Enabled = value;
            checkBoxSum.Enabled = value;
            buttonQuery.Enabled = value;
            buttonExcel.Enabled = value;
        }
        private async void buttonQuery_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(comboBoxClans.Text) /* && comboBoxClans.Text.Trim().Length >= TagLen */)
            {
                Cursor = Cursors.WaitCursor;
                reoGridControl.Cursor = Cursors.WaitCursor;
                setCheckBoxState(false);
                labelDate.Text = "查询时间：" + DateTime.Now.ToLongDateString();
                try
                {
                    Clan clan = await clashRoyale.GetClanAsync(comboBoxClans.Text.ToUpper());
                    if (clan != null)
                    {
                        SetGridHeader();
                        progressBarQuery.Maximum = clan.memberCount + 53;
                        progressBarQuery.Value = 0;

                        checkBoxType.Checked = (clan.type == "open");
                        comboBoxRegion.SelectedItem = clan.location.code;

                        textBoxDescription.Text = clan.description;
                        textBoxRequiredScore.Text = clan.requiredScore.ToString();
                        textBoxMemberCount.Text = clan.memberCount.ToString();
                        textBoxDonations.Text = clan.donations.ToString();
                        textBoxScore.Text = clan.score.ToString();
                        labelName.Text = clan.name + " 统计表";
                        pictureBoxBadge.Image = await GetImage(clan.badge.image);

                        if (checkBoxWar.Checked)
                        {
                            clanwar = await clashRoyale.GetClanWarAsync(comboBoxClans.Text.ToUpper());
                            if (clanwar != null) comboBoxState.SelectedIndex = GetState(clanwar.state);
                            progressBarQuery.Value = 1;
                        }
                        if (checkBoxRecord.Checked)
                        {
                            clanwars = await clashRoyale.GetPastClanWarsAsync(comboBoxClans.Text.ToUpper());
                            progressBarQuery.Value = 2;
                        }
                        progressBarQuery.Value = 3;

                        // 这里是最慢的地方，基本占用时间超过 50%
                        IEnumerable<string> playertags = from p in clan.members select p.tag;
                        players = await clashRoyale.GetPlayerAsync(playertags.ToArray(), checkBoxChest.Checked, true);
                        progressBarQuery.Value = 53;
                        if (players != null && players.Length > 0)
                        {
                            //for (int i = 1; i <= clan.members.Length; i++)
                            for (int i = 1; i <= players.Length; i++)
                            {
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.id] = i;                                      // 序号

                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.name] = clan.members[i - 1].name;               // 姓名
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.role] = GetRole(clan.members[i - 1].role);      // 职位                        
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.score] = clan.members[i - 1].trophies;           // 杯数
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.level] = clan.members[i - 1].expLevel;           // 大本营
                                if (checkBoxViewCardImage.Checked)
                                    reoGridControl.CurrentWorksheet[i, (int)GridHeader.arena] = new ImageCell(GetArenaImage(clan.members[i - 1].arena));
                                else
                                    reoGridControl.CurrentWorksheet[i, (int)GridHeader.arena] = GetArenaName(clan.members[i - 1].arena);
                                /*                                 
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.name] = players[i - 1].name;               // 姓名
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.role] = GetRole(players[i - 1].role);      // 职位                        
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.score] = players[i - 1].trophies;           // 杯数
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.level] = players[i - 1].expLevel;           // 大本营                         
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.arena] = GetArenaName(players[i - 1].arena);   
                                */
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.incards] = clan.members[i - 1].donationsReceived;  // 收到卡牌
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.outcards] = clan.members[i - 1].donations;          // 捐出卡牌
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.deltacards] = clan.members[i - 1].donationsDelta;     // 卡牌盈余
                                #region 获取 成员信息
                                //Player player = await clashRoyale.GetPlayerAsync(clan.members[i - 1].tag, checkBoxChest.Checked, true);
                                Player player = players[i - 1];
                                if (player != null)
                                {
                                    //reoGridControl.CurrentWorksheet[i,(int)GridHeader.arena] = new HyperlinkCell(player.deckLink);
                                    if (player.stats != null)
                                    {
                                        // 擅长卡牌
                                        if (checkBoxViewCardImage.Checked)
                                        {
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.favoriteCard] = new ImageCell(GetCardImage(player.stats.favoriteCard));
                                            //reoGridControl.CurrentWorksheet.SetRowsHeight(i, 1, (ushort)imageListCards.ImageSize.Height);
                                            reoGridControl.CurrentWorksheet.SetRowsHeight(i, 1, (ushort)imageListChests.ImageSize.Height);
                                        }
                                        else
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.favoriteCard] = GetCardName(player.stats.favoriteCard);
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.totaloutcards] = player.stats.totalDonations;       // 捐牌总牌
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.cards] = player.stats.cardsFound - cards.GetLength(0);           // 已有卡牌数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.wins3] = player.stats.threeCrownWins;       // 三冠胜利数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.clanCards] = player.stats.clanCardsCollected;   // 部落卡牌总数
                                    }
                                    if (player.games != null)
                                    {
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.wins] = player.games.wins;                 // 胜利场数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.losses] = player.games.losses;               // 失败场数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.draws] = player.games.draws;                // 平局场数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.clanwins] = player.games.warDayWins;           // 部落战胜利场数
                                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.winsPercent] = player.games.winsPercent; // 胜率               
                                    }
                                    #region 显示 未来的 宝箱
                                    if (checkBoxChest.Checked)
                                    {
                                        for (int j = 0; j < player.chestCycle.upcoming.Length; j++)
                                        {
                                            if (checkBoxViewCardImage.Checked)
                                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest1 + j] = new ImageCell(getChestImage(player.chestCycle.upcoming[j]));
                                            else
                                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest1 + j] = GetChestName(player.chestCycle.upcoming[j]);
                                        }
                                        if (player.chestCycle.giant > MaxChest)
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest9 + 1] = "第 " + player.chestCycle.giant + " 个";
                                        if (player.chestCycle.magical > MaxChest)
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest9 + 2] = "第 " + player.chestCycle.magical + " 个";
                                        if (player.chestCycle.superMagical > MaxChest)
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest9 + 3] = "第 " + player.chestCycle.superMagical + " 个";
                                        if (player.chestCycle.epic > MaxChest)
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest9 + 4] = "第 " + player.chestCycle.epic + " 个";
                                        if (player.chestCycle.legendary > MaxChest)
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest9 + 5] = "第 " + player.chestCycle.legendary + " 个";

                                    }
                                    #endregion
                                    #region 显示 当前战斗
                                    if (checkBoxWar.Checked && clanwar.participants != null)
                                    {
                                        Participant participant = clanwar.participants.Where(Participant => Participant.tag == player.tag).FirstOrDefault();
                                        if (!string.IsNullOrEmpty(participant.tag))
                                        {
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardbattles] = participant.battlesPlayed;
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardwins] = participant.wins;
                                            reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardearns] = participant.cardsEarned;
                                        }
                                    }
                                    #endregion
                                    #region 显示 9 次战斗历史
                                    if (checkBoxRecord.Checked && clanwars != null)
                                    {
                                        int sum1 = 0, sum2 = 0;
                                        for (int j = 0; j < clanwars.Length; j++)
                                        {
                                            if (clanwars[j].participants != null)
                                            {
                                                Participant participant = clanwars[j].participants.Where(Participant => Participant.tag == player.tag).FirstOrDefault();
                                                if (!string.IsNullOrEmpty(participant.tag))
                                                {
                                                    reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardwins + 3 * j + 3] = participant.battlesPlayed;
                                                    reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardwins + 3 * j + 2] = participant.wins;
                                                    reoGridControl.CurrentWorksheet[i, (int)GridHeader.cardwins + 3 * j + 1] = participant.cardsEarned;
                                                    sum1 += participant.wins;
                                                    sum2 += participant.battlesPlayed;
                                                }
                                            }
                                        }
                                        if (checkBoxSum.Checked)
                                        {
                                            if (sum2 > 0)       //有参战
                                            {
                                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.clanwin] = sum1;
                                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.clanwinpercent] = 1.0f * sum1 / sum2;
                                                if (sum1 <= 0)  //胜率为 0
                                                {
                                                    reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                            new WorksheetRangeStyle
                                                                            {
                                                                                Flag = PlainStyleFlag.TextColor,
                                                                                TextColor = Color.SaddleBrown,
                                                                            }));
                                                }
                                                else //胜率不为零
                                                {
                                                    reoGridControl.CurrentWorksheet.SetCellBody(i, (int)GridHeader.clanwinpercent, new NumericProgressCell());
                                                    if (sum1 >= clanwars.Length / 2 && sum2 >= clanwars.Length) //胜率一半以上，而且全部参加
                                                    {
                                                        reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                                new WorksheetRangeStyle
                                                                                {
                                                                                    Flag = PlainStyleFlag.TextColor,
                                                                                    TextColor = Color.Green,
                                                                                }));
                                                    }
                                                }
                                            }
                                            else// 没有参战
                                            {
                                                // 捐兵高于一半成员
                                                if (clan.members[i - 1].donations != null && clan.members[i - 1].donations >= clan.donations / clan.memberCount)
                                                {
                                                    reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                            new WorksheetRangeStyle
                                                                            {
                                                                                Flag = PlainStyleFlag.TextColor,
                                                                                TextColor = Color.Orange,
                                                                            }));
                                                }
                                                else // 捐兵也很少
                                                {
                                                    reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                            new WorksheetRangeStyle
                                                                            {
                                                                                Flag = PlainStyleFlag.TextColor,
                                                                                TextColor = Color.Red,
                                                                            }));
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region 特殊 显示 首领
                                    if (clan.members[i - 1].tag == "99VCU9JPG")     //作者 低调显示
                                    {
                                        reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                new WorksheetRangeStyle
                                                                {
                                                                    Flag = PlainStyleFlag.TextColor,
                                                                    TextColor = Color.LightGray,
                                                                }));
                                    }
                                    else if (clan.members[i - 1].role == "leader")  // 首领 特殊显示
                                    {
                                        reoGridControl.DoAction(new SetRangeStyleAction(i, 0, 1, (int)GridHeader.totalcolumns,
                                                                new WorksheetRangeStyle
                                                                {
                                                                    Flag = PlainStyleFlag.TextColor | PlainStyleFlag.FontStyleBold,
                                                                    TextColor = Color.Green,
                                                                    Bold = true,
                                                                }));
                                    }
                                    #endregion
                                    progressBarQuery.Value++;
                                }
                                #endregion
                            }
                        }
                        if (checkBoxSum.Checked) GetRowSum();
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Source.Trim() == "CR API csharp wrapper")
                        MessageBox.Show("没有找到该部落，TAG号<" + comboBoxClans.Text + "> 是正确的么？如果确信没错，请将该号码 eMail 给作者, 作者会尽快回复并解决！", "没有找到 " + comboBoxClans.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                        labelDate.Text = "错误" + ":" + ex.Message + "-" + ex.StackTrace;
                }
                finally
                {
                    setCheckBoxState(true);
                    reoGridControl.Cursor = Cursors.Default;
                    Cursor = Cursors.Default;
                }
            }
        }
        private void UpdateGrid(bool ShowImage)
        {
            Cursor = Cursors.WaitCursor;
            if (players != null && players.Length > 0)
            {
                if (checkBoxChest.Checked)
                    reoGridControl.CurrentWorksheet.ClearRangeContent(new RangePosition(1, (int)GridHeader.chest1, players.Length, MaxChest), CellElementFlag.Body | CellElementFlag.Data);
                reoGridControl.CurrentWorksheet.ClearRangeContent(new RangePosition(1, (int)GridHeader.favoriteCard, players.Length, 1), CellElementFlag.Body | CellElementFlag.Data);
                reoGridControl.CurrentWorksheet.ClearRangeContent(new RangePosition(1, (int)GridHeader.arena, players.Length, 1), CellElementFlag.Body | CellElementFlag.Data);

                if (ShowImage)
                {
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.chest1, 9, (ushort)imageListChests.ImageSize.Width);  // * imageListCards.ImageSize.Height / imageListChests.ImageSize.Height));
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.favoriteCard, 1, (ushort)(imageListCards.ImageSize.Width * imageListChests.ImageSize.Height / imageListCards.ImageSize.Height));
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.arena, 1, (ushort)(imageListArena.ImageSize.Width * imageListChests.ImageSize.Height / imageListArena.ImageSize.Height));
                    for (int i = 1; i <= players.Length; i++)
                    {
                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.favoriteCard] = new ImageCell(GetCardImage(players[i - 1].stats.favoriteCard));
                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.arena] = new ImageCell(GetArenaImage(players[i - 1].arena));
                        if (checkBoxChest.Checked)
                        {
                            for (int j = 0; j < players[i - 1].chestCycle.upcoming.Length; j++)
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest1 + j] = new ImageCell(getChestImage(players[i - 1].chestCycle.upcoming[j]));
                        }
                    }
                }
                else
                {
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.chest1, 9, HeaderWidth[(int)GridHeader.chest1]);
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.favoriteCard, 1, HeaderWidth[(int)GridHeader.favoriteCard]);
                    reoGridControl.CurrentWorksheet.SetColumnsWidth((int)GridHeader.arena, 1, HeaderWidth[(int)GridHeader.arena]);
                    for (int i = 1; i <= players.Length; i++)
                    {
                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.favoriteCard] = GetCardName(players[i - 1].stats.favoriteCard);
                        reoGridControl.CurrentWorksheet[i, (int)GridHeader.arena] = GetArenaName(players[i - 1].arena);
                        if (checkBoxChest.Checked)
                        {
                            for (int j = 0; j < players[i - 1].chestCycle.upcoming.Length; j++)
                                reoGridControl.CurrentWorksheet[i, (int)GridHeader.chest1 + j] = GetChestName(players[i - 1].chestCycle.upcoming[j]);
                        }
                    }
                }
            }
            Cursor = Cursors.Default;
        }
        private void GetRowSum()
        {
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.score] = "=SUM(" + new RangePosition(1, (int)GridHeader.score, MaxMember, 1).ToRelativeAddress() + ")";
            //reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.level] = "=AVERAGE(" + new RangePosition(1, (int)GridHeader.level, MaxMember, 1).ToRelativeAddress() + ")";
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.incards] = "=SUM(" + new RangePosition(1, (int)GridHeader.incards, MaxMember, 1).ToRelativeAddress() + ")";
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.cards] = "=SUM(" + new RangePosition(1, (int)GridHeader.cards, MaxMember, 1).ToRelativeAddress() + ")";
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.clanwins] = "=SUM(" + new RangePosition(1, (int)GridHeader.clanwins, MaxMember, 1).ToRelativeAddress() + ")";
            reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.winsPercent] = "=AVERAGE(" + new RangePosition(1, (int)GridHeader.winsPercent, MaxMember, 1).ToRelativeAddress() + ")";
            reoGridControl.CurrentWorksheet.SetCellBody(SumRow, (int)GridHeader.winsPercent, new NumericProgressCell());
            if (checkBoxWar.Checked)
            {
                reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.cardbattles] = "=SUM(" + new RangePosition(1, (int)GridHeader.cardbattles, MaxMember, 1).ToRelativeAddress() + ")";
                reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.cardwins] = "=SUM(" + new RangePosition(1, (int)GridHeader.cardwins, MaxMember, 1).ToRelativeAddress() + ")";
                reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.cardearns] = "=SUM(" + new RangePosition(1, (int)GridHeader.cardearns, MaxMember, 1).ToRelativeAddress() + ")";
            }
            if (checkBoxRecord.Checked)
            {
                reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.clanwin] = "=SUM(" + new RangePosition(1, (int)GridHeader.clanwin, MaxMember, 1).ToRelativeAddress() + ")";
                reoGridControl.CurrentWorksheet[SumRow, (int)GridHeader.clanwinpercent] = "=AVERAGE(" + new RangePosition(1, (int)GridHeader.clanwinpercent, MaxMember, 1).ToRelativeAddress() + ")";
                reoGridControl.CurrentWorksheet.SetCellBody(SumRow, (int)GridHeader.clanwinpercent, new NumericProgressCell());
            }
        }
        private string GetRole(string Role)
        {
            if (!string.IsNullOrEmpty(Role))
            {
                string role = Role.ToLower().Trim();
                if (role == "leader") return "首领";
                else if (role == "coleader") return "副首领";
                else if (role == "elder") return "长老";
                else if (role == "member") return "成员";
                else return role;
            }
            return "";
        }
        private string GetChestName(string Chest)
        {
            if (!string.IsNullOrEmpty(Chest))
            {
                string chest = Chest.ToLower().Trim();
                if (chest == "epic") return chests[7];
                else if (chest == "silver") return chests[1];
                else if (chest == "gold") return chests[2];
                else if (chest == "giant") return chests[3];
                else if (chest == "magical") return chests[4];
                else if (chest == "super magical") return chests[5];
                else if (chest == "legendary") return chests[6];
                else return chest;
            }
            return "";
        }
        private Image getChestImage(string Chest)
        {
            if (!string.IsNullOrEmpty(Chest))
            {
                if (imageListChests.Images.ContainsKey(Chest))
                    return imageListChests.Images[Chest];
            }
            return null;
        }
        private int GetState(string State)
        {
            if (!string.IsNullOrEmpty(State))
            {
                string state = State.ToLower().Trim();
                if (state == "collectionday")
                {
                    pictureBox5.Image = Resources.cards;
                    labelEndTime.Text = "结束：" + IntToDateTime(clanwar.stageEndTime).ToString();
                    return 1;
                }
                else if (state == "warday")
                {
                    pictureBox5.Image = Resources.battle;
                    labelEndTime.Text = "结束：" + IntToDateTime(clanwar.stageEndTime).ToString();
                    return 2;
                }
            }
            pictureBox5.Image = null;
            return 0;
        }
        private string GetCardName(Card card)
        {
            if (card != null && !string.IsNullOrEmpty(card.key))
            {
                for (int i = 0; i < cards.GetLength(0); i++)
                {
                    if (card.key == cards[i, 0]) return cards[i, 2];
                }
                return card.name;
            }
            return "无";
        }
        private Image GetCardImage(Card card)
        {
            if (card != null && !string.IsNullOrEmpty(card.key))
            {
                if (imageListCards.Images.ContainsKey(card.key))
                    return imageListCards.Images[card.key];
            }
            return null;
        }
        private Image GetArenaImage(Arena arena)
        {
            if (arena != null && arena.arenaID > 0 && arena.arenaID <= imageListArena.Images.Count)
                return imageListArena.Images[arena.arenaID - 1];
            return null;
        }
        private string GetArenaName(Arena arena)
        {
            if (arena != null)
            {
                if (arena.arenaID >= 0 && arena.arenaID < arenas.Length)
                    return arena.arenaID + " : " + arenas[arena.arenaID];
                else
                    return arena.arenaID + " : " + arena.name;
            }
            else
                return "无";
        }
        private DateTime IntToDateTime(object utc)
        {
            try
            {
                DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
                startTime = startTime.AddSeconds(Convert.ToInt32(utc));
                return startTime;
            }
            catch (Exception)
            {
                return DateTime.Now;
            }
        }
        private Task<Image> GetImage(string uri)
        {
            return Task.Run(() =>
            {
                try
                {
                    //处理HttpWebRequest访问https有安全证书的问题（ 请求被中止: 未能创建 SSL/TLS 安全通道。）
                    ServicePointManager.ServerCertificateValidationCallback += (s, cert, chain, sslPolicyErrors) => true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    Image img = Image.FromStream(WebRequest.Create(uri).GetResponse().GetResponseStream());
                    return img;
                }
                catch (WebException ex)
                {
                    labelDate.Text = "错误" + ":" + ex.Message + "-" + ex.StackTrace;
                    return null;
                }
            });
        }
        private void buttonAbout_Click(object sender, EventArgs e)
        {
            AboutBox frm = new AboutBox();
            if (DialogResult.Cancel == frm.ShowDialog())
            {
                comboBoxClans.Text = "9P88QL8C";
            }
        }

        private void checkBoxDatabase_CheckedChanged(object sender, EventArgs e)
        {
            //buttonSave.Enabled = checkBoxDatabase.Checked;
        }

        private async void buttonBestClans_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            comboBoxClans.Enabled = false;
            buttonBestClans.Enabled = false;
            labelDate.Text = "查询时间：" + DateTime.Now.ToLongDateString();
            try
            {
                SimplifiedClan[] bestClans = null;
                if (checkBoxRegion.Checked)
                    bestClans = await clashRoyale.GetTopClansAsync(GameLocation);
                else
                    bestClans = await clashRoyale.GetTopClansAsync();
                if (bestClans != null && bestClans.Length > 0)
                {
                    comboBoxClans.Items.Clear();
                    foreach (SimplifiedClan clan in bestClans)
                        comboBoxClans.Items.Add(clan.tag);
                }
            }
            catch (Exception ex)
            {
                labelDate.Text = "错误" + ":" + ex.Message + "-" + ex.StackTrace;
            }
            finally
            {
                comboBoxClans.Enabled = true;
                buttonBestClans.Enabled = true;
                Cursor = Cursors.Default;
            }
        }
        private Wrapper.Locations GameLocation { get { return Wrapper.Locations._EU + comboBoxRegion.SelectedIndex; } }

        private void comboBoxRegion_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxRegion.SelectedIndex >= 0) checkBoxRegion.Enabled = true;
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBoxClans.Text))
            {
                string file = Application.ExecutablePath;
                Configuration config = ConfigurationManager.OpenExeConfiguration(file);
                if (config.AppSettings.Settings.AllKeys.Contains(TagName))
                    config.AppSettings.Settings.Remove(TagName);
                config.AppSettings.Settings.Add(TagName, comboBoxClans.Text);
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            reoGridControl.Save(
                Path.Combine(Application.StartupPath, comboBoxClans.Text + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls"),
                unvell.ReoGrid.IO.FileFormat.Excel2007);
        }

        private void checkBoxViewCardImage_CheckedChanged(object sender, EventArgs e)
        {
            UpdateGrid(checkBoxViewCardImage.Checked);
        }
    };

    internal class NumericProgressCell : CellBody
    {
        public override void OnPaint(CellDrawingContext dc)
        {
            float value = 0;
            float.TryParse(dc.Cell.Data.ToString(), out value);
            if (value > 0)
            {
                IGraphics g = dc.Graphics;
                unvell.ReoGrid.Graphics.Rectangle rect = new unvell.ReoGrid.Graphics.Rectangle(Bounds.Left, Bounds.Top + 1, (int)(Math.Round(value * Bounds.Width)), Bounds.Height - 1);
                g.FillRectangleLinear(Color.Coral, Color.IndianRed, 0f, rect);
                g.DrawText(dc.Cell.DisplayText, "黑体", 10f, SolidColor.DarkBlue, Bounds);
            }
        }
    };
}
