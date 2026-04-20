import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from "docx";
import { writeFileSync } from "fs";

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "游明朝", size: 24 },
      },
    },
  },
  sections: [
    {
      children: [

        // タイトル
        new Paragraph({
          text: "子どもに教えられるお母さんになる　AI時代の家族を守る思考法",
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // 目次
        new Paragraph({ text: "目次", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
        new Paragraph({ text: "はじめに　「お母さん、AIって何？」——その質問に、あなたは答えられましたか？" }),
        new Paragraph({ text: "おわりに　あなたが変われば、家族が変わる。家族が変われば、未来が変わる。" }),
        new Paragraph({ text: "第1章　知らないうちに、子どもの「考える力」が奪われている" }),
        new Paragraph({ text: "第2章　「あのお母さん、なんか違う」と思われる存在になる" }),
        new Paragraph({ text: "第3章　難しくない。5つの問いかけが家族を変える" }),
        new Paragraph({ text: "第4章　データと実例が証明する「考える力」の育て方" }),
        new Paragraph({ text: "第5章　今日からできる「家族を守る思考法」完全ガイド" }),
        new Paragraph({ text: "第6章　明日の朝、子どもにこの一言を言ってみてください", spacing: { after: 400 } }),

        // ── はじめに ──
        new Paragraph({ text: "はじめに｜「お母さん、AIって何？」——その質問に、あなたは答えられましたか？", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),
        new Paragraph({ text: "ある夜のことです。小学4年生の息子がリビングでタブレットを触りながら、突然こう言いました。「ねえお母さん、AIって結局何なの？学校の先生もよくわかってないみたいだったんだけど」", spacing: { after: 200 } }),
        new Paragraph({ text: "そのお母さんは、言葉に詰まったと言います。「え……なんて答えればいいんだろう」スマホを取り出してGoogle検索しようとしたとき、ふと手が止まりました。「私、子どもの質問にいつもスマホで調べてから答えてる。自分の言葉で答えたこと、最近あったかな……」", spacing: { after: 200 } }),
        new Paragraph({ text: "この本は、そんな瞬間に心当たりのあるお母さんに向けて書きました。", spacing: { after: 300 } }),
        new Paragraph({ text: "AIが当たり前になった時代に、私たちは何を失いかけているのか。", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "ChatGPT、Gemini、Copilot——気づけばスマホの中にAIがあふれています。献立を聞けば答えてくれる。子どもの宿題も調べてくれる。悩みを打ち明ければ優しく返してくれる。便利です。本当に、便利です。", spacing: { after: 200 } }),
        new Paragraph({ text: "でもその便利さの影で、じわじわと失われていくものがあります。「自分で考える力」です。大人だけじゃない。子どもも同じです。答えがすぐ手に入る環境で育った子どもは、答えの出ない問いに直面したとき、どうすればいいかわからなくなる。悩む前にAIを開く。考える前に検索する。その繰り返しの中で、「自分はどう思うか」という感覚が少しずつ薄れていく。", spacing: { after: 200 } }),
        new Paragraph({ text: "これは脅かしではありません。今、学校現場でも、育児の現場でも、静かに起きていることです。", spacing: { after: 300 } }),
        new Paragraph({ text: "でも、安心してください。", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "この本はAIを否定しません。スマホを捨てろとも、検索するなとも言いません。お伝えしたいのはたった一つのことです。AIを「使う人間」の側に、考える軸を持つこと。そしてその軸を、お母さん自身が持ち、子どもに手渡していくこと。", spacing: { after: 200 } }),
        new Paragraph({ text: "難しいスキルは要りません。資格も要りません。必要なのは、日常会話の中に「5つの問いかけ」を混ぜるだけ。たったそれだけで、子どもの思考習慣は変わります。家族の会話が変わります。そして気づいたら、あなた自身も変わっています。", spacing: { after: 200 } }),
        new Paragraph({ text: "この本を読み終えたとき、きっとこう思えるはずです。「明日の朝、子どもにあの一言を言ってみよう」と。その一言から、全てが始まります。さあ、一緒に始めましょう。", spacing: { after: 400 } }),

        // ── 第1章 ──
        new Paragraph({ text: "第1章｜知らないうちに、子どもの「考える力」が奪われている", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "突然ですが、こんな場面に心当たりはありませんか？", spacing: { after: 200 } }),
        new Paragraph({ text: "夕食の後、小学3年生の娘が宿題のプリントを持ってきました。「お母さん、これわからない」あなたは一瞬考えて……スマホを取り出し、ChatGPTに質問を打ち込みます。答えはすぐに出てきました。娘はそれをノートに書き写して「ありがとう！」と笑顔で去っていきます。", spacing: { after: 200 } }),
        new Paragraph({ text: "問題は解決した。でも、なんだか胸にモヤが残る。", spacing: { after: 200 } }),
        new Paragraph({ text: "「これって、本当によかったのかな……」", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：AIは便利だからこそ、使い方を間違えると「考える力」を静かに奪っていく", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "はっきり言います。AIそのものは悪くありません。使い方次第で、最高の道具にも、最悪の思考停止装置にもなります。", spacing: { after: 200 } }),
        new Paragraph({ text: "問題は、「答えを出すこと」と「考えること」を混同してしまうことです。AIはどんな質問にも即座に答えを返してくれます。だからこそ、「自分で考える前にAIに聞く」習慣がついてしまうと、考えるプロセスごとAIに丸投げすることになる。そしてそれが毎日続くと、気づかないうちに「自分で考える筋肉」が少しずつ落ちていくんです。", spacing: { after: 200 } }),
        new Paragraph({ text: "これは大人だけじゃなく、子どもにも同じことが起きています。", spacing: { after: 300 } }),

        new Paragraph({ text: "理由：子どもの脳は「今」つくられている", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "小学生の時期は、脳の発達において非常に重要な時期です。「なぜだろう？」「どうしたらいいんだろう？」と自分なりに悩み、試行錯誤する経験の積み重ねが、考える力の土台をつくります。", spacing: { after: 200 } }),
        new Paragraph({ text: "ところが、悩む前にAIが答えを出してしまうと、その「悩む時間」ごと消えてしまう。答えは手に入る。でも、考えるプロセスは育たない。", spacing: { after: 200 } }),
        new Paragraph({ text: "子どもにとって宿題は「答えを出すこと」が目的ではなく、「考えること」が目的のはずです。なのに、AIを使うことで、その本来の目的が知らないうちにすり替わってしまっているんです。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例①：よくある「失敗パターン」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "横浜に住む38歳の主婦・Aさんの話をご紹介します。Aさんは子ども思いの、とても真面目なお母さんです。子どもが「わからない」と言うたびに、スマホでAIに質問して答えを教えていました。「子どもに困ってほしくない」という愛情からです。", spacing: { after: 200 } }),
        new Paragraph({ text: "ところが半年後、小学4年生の息子に変化が現れました。算数の文章題を前にしてすぐ「わからない、調べていい？」と言うようになった。自分で考えようとする時間が、どんどん短くなっていったんです。", spacing: { after: 200 } }),
        new Paragraph({ text: "担任の先生から「授業中、少し考える前にすぐ手が止まってしまうことがあります」と言われて、Aさんははじめて気づきました。「私が、子どもから考える機会を奪っていたのかもしれない」", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例②：変化した「成功パターン」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "Aさんはそこから少し変えました。AIに答えを聞く前に、まず子どもに1つだけ質問するようにしたんです。「どこまではわかった？」", spacing: { after: 200 } }),
        new Paragraph({ text: "たったこれだけです。すると息子は「ここまではわかるんだけど、ここからがわからない」と自分で整理して話し始めました。その「整理する時間」こそが、考える力を育てるプロセスだったんです。", spacing: { after: 200 } }),
        new Paragraph({ text: "1ヶ月後、息子は「お母さん、ちょっと待って、自分でやってみる」と言うようになりました。Aさんは「たった一言の問いがこんなに変わるとは思わなかった」と話しています。", spacing: { after: 200 } }),
        new Paragraph({ text: "AIを使うことをやめたわけじゃない。ただ、使う順番を変えただけです。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：お母さんが「考え方」を持てば、家族が変わる", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "大切なのは、AIを使わせないことではありません。AIと一緒に考える力を育てることです。そのためにお母さん自身が「AI時代の思考法」を持つことが、何よりの近道です。難しいスキルも、特別な知識も必要ありません。必要なのは、ちょっとした「問いかけの習慣」だけ。", spacing: { after: 200 } }),
        new Paragraph({ text: "次の章では、そんなお母さんたちが実際にどう変わったのか、具体的な姿をお伝えします。あなたにも、きっとできます。", spacing: { after: 400 } }),

        // ── 第2章 ──
        new Paragraph({ text: "第2章｜「あのお母さん、なんか違う」と思われる存在になる", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "あなたの周りに、こんなお母さんはいませんか？", spacing: { after: 200 } }),
        new Paragraph({ text: "PTAの集まりで、AIの話題になったとき。「うちの子、最近ChatGPTで宿題してるんだけど、どう思う？」という話に、みんなが「うーん、どうなんだろうね」と曖昧に笑う中、一人だけ落ち着いてこう言うお母さん。", spacing: { after: 200 } }),
        new Paragraph({ text: "「うちはね、AIを使う前に『自分はどう思う？』って一回聞くようにしてるの。そうすると子どもが自分の言葉で話してくれるようになってきたよ」", spacing: { after: 200 } }),
        new Paragraph({ text: "場がしんとして、誰かが「それいいね、どうやるの？」と前のめりになる。そのお母さんは特別な資格を持っているわけでも、ITの仕事をしているわけでもない。ただ、「考え方の軸」を持っているだけです。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：憧れのお母さんは、知識量が多いのではなく「問いの立て方」が違う", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "はっきり言います。AI時代に「頼りになるお母さん」になるために、難しい技術を学ぶ必要はありません。必要なのは知識の量ではなく、「なぜ？」「本当に？」「他には？」という3つの問いを持つ習慣です。", spacing: { after: 200 } }),
        new Paragraph({ text: "この習慣を持つお母さんは、子どもの目に「なんか違う」と映ります。それは威圧感ではなく、信頼感です。「このお母さんに話すと、自分の考えが整理される」という安心感です。それこそが、AI時代に家族の軸になれるお母さんの姿です。", spacing: { after: 300 } }),

        new Paragraph({ text: "理由：子どもが本当に求めているのは「答え」ではなく「一緒に考えてくれる人」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "子どもが「わからない」と言うとき、実は答えだけを求めているわけではありません。「一緒に考えてほしい」「自分の気持ちをわかってほしい」という気持ちが、その言葉の奥にあることがほとんどです。", spacing: { after: 200 } }),
        new Paragraph({ text: "AIは瞬時に答えを返してくれます。でも「一緒に考えてくれる」ことはできません。共感も、励ましも、「あなたならできる」という眼差しも、AIには出せない。だからこそ、「考える過程に寄り添えるお母さん」は、AIには絶対に代替できない存在です。これはAI時代だからこそ、より輝く強みなんです。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例①：変わる前のBさん", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "埼玉に住む40歳の主婦・Bさんは、自分に自信がありませんでした。高卒で、ITの知識もない。「AIとかよくわからないし、子どもに何か聞かれても答えられないな」と感じていました。", spacing: { after: 200 } }),
        new Paragraph({ text: "子どもがAIで調べた答えをそのまま宿題に書いていても、「合ってるならいいか」と止められなかった。正しいかどうか判断できる自信がなかったから。夫に「お母さん、AIわからないの？」と子どもに笑われた日、Bさんは布団の中で小さく泣いたと言います。「私って、何も教えられないお母さんなんだ」", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例②：変わった後のBさん", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "Bさんが変わったきっかけは、たった一つの習慣を始めたことでした。子どもがAIの答えを見せてきたとき、こう聞くようにしたんです。「へえ、面白いね。でもこれって、本当にそうなのかな？」", spacing: { after: 200 } }),
        new Paragraph({ text: "答えを否定するわけじゃない。ただ、一緒に「本当かな？」と疑う。最初は子どもも「え、違うの？」と戸惑っていました。でも2週間もすると、子ども自身が「これ本当かな、調べてみよう」と言うようになった。", spacing: { after: 200 } }),
        new Paragraph({ text: "そしてある日、子どもがこう言ったんです。「お母さん、AIって嘘つくこともあるんだね」Bさんは胸が熱くなったと話します。「私が何かを教えたんじゃなくて、一緒に気づけた気がして。それが嬉しかった」", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：あなたには、すでに武器がある", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "30代・40代のお母さんには、AIにはない最強の武器があります。それは「人生経験」です。失敗したこと、悩んだこと、遠回りしてきたこと。その全てが「問いを立てる力」の土台になります。", spacing: { after: 200 } }),
        new Paragraph({ text: "AIは過去のデータから答えを出します。でもあなたは、自分の経験から「本当のこと」を知っている。その感覚を、思考法という形で子どもに手渡せるのは、あなただけです。次の章では、その「思考法」を具体的な5つの問いとして、誰でもすぐに使える形でお伝えします。「あのお母さん、なんか違う」——それはあなたの話です。", spacing: { after: 400 } }),

        // ── 第3章 ──
        new Paragraph({ text: "第3章｜難しくない。5つの問いかけが家族を変える", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "「思考法」と聞いて、こう思いませんでしたか？「なんか難しそう……」「哲学とか、頭のいい人がやることでしょ？」「毎日忙しいのに、そんな余裕ないよ」", spacing: { after: 200 } }),
        new Paragraph({ text: "わかります。私も最初そう思いました。でも安心してください。ここでお伝えする思考法は、哲学書に書いてあるような難しい話ではありません。夕飯の支度をしながらでも、子どもの送り迎えの車の中でも、今日からすぐに使える「5つの問いかけ」です。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：AI時代の思考法は「5つの問い」を口グセにするだけでいい", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "① 「それって、誰が言ってたの？」\n② 「本当にそうかな？他の見方はないかな？」\n③ 「もし違ったら、どうなると思う？」\n④ 「あなたはどう思う？」\n⑤ 「それを知って、どうしたい？」", spacing: { after: 200 } }),
        new Paragraph({ text: "この5つです。暗記しなくていい。全部使わなくていい。1日1つ、思い出したときに使うだけで十分です。", spacing: { after: 300 } }),

        new Paragraph({ text: "理由：思考力は「考える機会」の積み重ねで育つ", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "思考力は才能ではありません。筋肉と同じで、使えば育ち、使わなければ落ちていきます。AIが普及した今、「考える機会」は意識しないと自然に減っていきます。答えはすぐ手に入る。調べればわかる。悩む前に解決できる。これは便利である反面、考える筋肉を使う機会が減るということでもあります。", spacing: { after: 200 } }),
        new Paragraph({ text: "だからこそ、日常会話の中に「考えるきっかけ」を意図的に作ることが大切です。お母さんが問いを投げかけるだけで、子どもは「考える時間」を自然に持てるようになります。答えを教えなくていい。一緒に「うーん」と悩むだけでいい。その時間こそが、思考力を育てる黄金の時間です。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例①：5つの問いを使った「ある朝の会話」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "小学2年生の娘を持つCさん（36歳）の実例です。ある朝、娘がスマホでニュースを見ながら「ねえお母さん、AIって人間の仕事を全部なくすんだって！」と言いました。", spacing: { after: 200 } }),
        new Paragraph({ text: "Cさんは①の問いを使いました。「へえ、それって誰が言ってたの？」娘は「えっと……YouTubeで見た」と答えました。「そっか。じゃあ本当にそうかな？他の人はどう言ってるんだろうね」（②）", spacing: { after: 200 } }),
        new Paragraph({ text: "娘は少し考えてから「調べてみる！」と言ってスマホを手に取りました。そして「あ、なくなる仕事もあるけど、新しい仕事も増えるって書いてある」と自分で発見したんです。Cさんは答えを教えていません。ただ2つの問いを投げかけただけ。「これが思考法か、と初めて実感した瞬間でした」とCさんは言います。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例②：失敗から学んだDさんのケース", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "42歳のDさんは、5つの問いを覚えた翌日、張り切りすぎてしまいました。子どもが話すたびに矢継ぎ早に問いを投げかけてしまい、子どもは「もうお母さんに話したくない」と部屋に入ってしまいました。", spacing: { after: 200 } }),
        new Paragraph({ text: "Dさんは翌日こう反省したと言います。「問いかけは尋問じゃない。会話のスパイスだったんだ」1日1問に絞り、まず子どもの話をしっかり聞くことを優先するようにしたら、1週間後には子どもが自分から話しかけてくるようになったそうです。使いすぎない。それが唯一のコツです。", spacing: { after: 400 } }),

        new Paragraph({ text: "結論：5つの問いは、家族の会話を変える魔法のスパイス", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "難しい知識は要りません。特別な資格も要りません。必要なのは、5つの問いを「口グセ」にする小さな勇気だけです。今日の夕飯のとき、子どもが何か話したら、ひとつだけ使ってみてください。「それって、誰が言ってたの？」たったその一言が、あなたと子どもの関係を、そして家族の未来を、少しずつ変えていきます。", spacing: { after: 400 } }),

        // ── 第4章 ──
        new Paragraph({ text: "第4章｜データと実例が証明する「考える力」の育て方", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "「わかった。でも、これって本当に効果があるの？」そう思うのは当然です。この章ではデータと専門家の言葉、そして実際の家庭の変化をもとに、しっかりお伝えします。感覚じゃなく、事実で確認してください。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：データは明確に示している。AI時代に必要なのは「問いを立てる力」だ", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "日本人の7割以上がAIに対して不安を感じているというデータがあります。そしてその不安の中身を見ると、「仕事がなくなる」という経済的な不安よりも、「自分で考えられなくなるのでは」という思考力への不安が年々増加していることがわかっています。", spacing: { after: 200 } }),
        new Paragraph({ text: "さらに女性は男性と比べて「AIによる思考力・創造性の低下」への不安が有意に高い。つまり、あなたが感じているその不安は、データに裏付けられた正しい直感なのです。", spacing: { after: 300 } }),

        new Paragraph({ text: "理由：AIは「答えを出す」が、「考えるプロセス」は代替できない", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "専門家たちは口をそろえてこう言います。「AIが進化するほど、人間に求められるのは答えを出す力ではなく、問いを立てる力だ」多くのAI研究者が指摘するのは「AIは過去のデータから最適解を導く。しかし、問題そのものを発見し、問いを立てるのは人間にしかできない」という点です。", spacing: { after: 200 } }),
        new Paragraph({ text: "これからの時代に必要な力は、より多くの情報を知ることではなく、情報に対して「本当に？」「なぜ？」と問いを立てられること。そしてその力は、子ども時代の日常的な「問いかけ習慣」によって育まれます。学校の授業でも、塾でもなく、家庭の会話の中で育つ力なんです。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例①：「AI依存」が引き起こした実際のリスク", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "精神科医・臨床心理士たちが最近警告しているのが、「AI依存」の問題です。AIに何でも相談するようになった人が、自分の判断に自信が持てなくなるケースが増えています。AIが「こうすべきです」と言うと、自分の感覚より優先してしまう。繰り返すうちに「自分はどう思うか」がわからなくなっていく。", spacing: { after: 200 } }),
        new Paragraph({ text: "ある調査では、小学生の4割以上が「わからないことがあったらまずAIやスマホで調べる」と回答。「自分で少し考えてから調べる」と答えた子どもは年々減っています。これは子どもが悪いわけじゃない。便利なものが目の前にあれば、そちらに頼るのは自然なことです。だからこそ、「考える前に一呼吸置く習慣」を意図的に作る必要があるのです。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例②：思考力を育てた家庭に共通する「たった一つの行動」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "教育研究者たちが、子どもの思考力が高い家庭を調査したところ、共通する行動が一つ見つかりました。それは、「親が答えを教えるより先に、子どもに問いを返す」習慣です。「答えを一緒に考える親」の子どもは、自分で考えることへの抵抗が低く、失敗を恐れずに試行錯誤できる傾向が強いことがわかっています。", spacing: { after: 200 } }),
        new Paragraph({ text: "逆に「すぐに正解を教える親」の子どもは、答えの出ない問いに直面したとき、極端に不安になりやすいというデータもあります。AIは答えを出してくれるけど、答えのない問いに立ち向かう力はAIには育てられない。それを育てられるのは、そばで「あなたはどう思う？」と問いを返してくれる親だけです。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：データが証明する、お母さんの「問いかけ」の力", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "・日本人の7割以上がAIへの不安を持ち、特に女性は思考力低下への懸念が強い\n・子どもの思考力を育てるのは授業でも塾でもなく、家庭の会話習慣\n・「問いを返す親」を持つ子どもは、考えることへの抵抗が低く、自立的に育つ\n・AI依存を防ぐ最大の対策は、日常の中に「考える一呼吸」を作ること", spacing: { after: 200 } }),
        new Paragraph({ text: "これは全て、あなたが今日からできることです。次の章では、ここまでの内容を実際の行動に落とし込んだ「家族を守る思考法 完全ガイド」をお伝えします。", spacing: { after: 400 } }),

        // ── 第5章 ──
        new Paragraph({ text: "第5章｜今日からできる「家族を守る思考法」完全ガイド", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "「わかった。でも、具体的に何から始めればいいの？」この章では、頭でわかったことを「実際の行動」に変えるための3つのステップをお伝えします。順番通りに進めてください。焦らなくていい。1週間ずつ、丁寧にやれば十分です。", spacing: { after: 300 } }),

        new Paragraph({ text: "STEP1：まずお母さん自身の思考を取り戻す", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "最初の1週間は、子どもに何かを伝えようとしなくていいです。まず自分自身のためにやってみてください。やること：毎日1つだけ「自分の答え」を出してみる。朝のニュース、SNSで流れてきた情報、ママ友の話。何でもいいので、AIやスマホで調べる前に30秒だけ「私はどう思う？」と自分に聞いてみる習慣をつけます。", spacing: { after: 200 } }),
        new Paragraph({ text: "Eさん（39歳）はこれを始めた翌週、「久しぶりに自分の意見を夫に話せた気がする」と言いました。特別なことは何もしていない。ただ30秒、自分に問いかけただけです。まずここから始めてください。", spacing: { after: 300 } }),

        new Paragraph({ text: "STEP2：子どもへの伝え方——年齢別の言葉の選び方", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "【小学校低学年（1〜3年生）向け】\nこの年齢に有効なのは「どう思う？」の一言を会話の最後に添えること。「今日学校どうだった？」→「楽しかった！」→「どのへんが一番楽しかった？」この3語を口グセにするだけで、子どもは自分の気持ちを言語化する力が育ちます。", spacing: { after: 200 } }),
        new Paragraph({ text: "【小学校高学年（4〜6年生）向け】\n有効なのは「それ、本当かな？一緒に確かめてみようか」という共同作業です。子どもが「〇〇らしいよ！」と言ったとき、否定も肯定もせず一緒に調べる。「ここではこう書いてるけど、こっちでは違うこと言ってるね。どっちが正しいと思う？」と問いかける。", spacing: { after: 200 } }),
        new Paragraph({ text: "Fさん（41歳）の小学5年生の息子は、この習慣を始めて1ヶ月後、学校の授業で「先生、それって本当ですか？」と手を挙げたそうです。先生から「とてもいい質問です」と褒められ、息子は誇らしげに帰ってきたと言います。", spacing: { after: 300 } }),

        new Paragraph({ text: "STEP3：家族で実践するAI時代のルールづくり", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "おすすめは「AIに聞く前に、30秒だけ自分で考える」ルールです。大事なのは、お母さんが一方的に決めないこと。夕食の後など、家族が集まったタイミングで子どもと一緒に決めることが大切です。「最近AIって便利だよね。でも使い方で差がつくらしいよ。うちはどうしようか？」", spacing: { after: 200 } }),
        new Paragraph({ text: "Gさん（37歳）の家族は「AIの答えをそのまま信じない。必ず一回『本当かな？』って言う」というルールを子どもと一緒に決めました。今では息子がお父さんに「それ、本当かな？」とツッコミを入れるようになり、家族の笑いのネタになっているそうです。", spacing: { after: 300 } }),

        new Paragraph({ text: "失敗しないための「3つの注意点」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "①完璧を目指さない：毎日できなくても大丈夫。週3回でも十分です。\n②結果をすぐ求めない：思考習慣の変化は3週間以上かかります。まず3週間続けることだけを目標に。\n③子どもを試さない：「ちゃんと考えてる？」は逆効果。「一緒に考えよう」というスタンスを忘れずに。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：今日の夜から始められる、3つのアクション", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "今夜：自分に「私はどう思う？」を1回だけ聞いてみる\n今週：子どもの話に「どのへんが？」を1回添えてみる\n今月：家族と「AI前の30秒ルール」を一緒に決めてみる", spacing: { after: 200 } }),
        new Paragraph({ text: "これだけです。難しいことは何もない。今夜の夕食から、始められます。", spacing: { after: 400 } }),

        // ── 第6章 ──
        new Paragraph({ text: "第6章｜明日の朝、子どもにこの一言を言ってみてください", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),

        new Paragraph({ text: "ここまで読んでくれて、ありがとうございます。第1章から読んできたあなたは、もう気づいているはずです。AI時代の思考法は、難しくない。特別な知識も、資格も、時間も要らない。必要なのは、「一言の問いかけ」を日常に混ぜる勇気だけだと。", spacing: { after: 300 } }),

        new Paragraph({ text: "結論：「考える力」を守れるのは、お母さんだけです", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "学校の先生でも、塾でも、AIでもなく。毎日そばにいるお母さんだけが、子どもの思考習慣を育てられる。これは責任ではなく、特権です。誰も代わりにできない、あなただけの役割です。", spacing: { after: 300 } }),

        new Paragraph({ text: "理由：習慣は「環境」が9割。家庭という環境が最強の学び場", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "教育心理学の研究では、子どもの学習習慣の形成に最も影響するのは学校でも塾でもなく、家庭環境と親の言動であることが繰り返し示されています。子どもは親を見て育ちます。お母さんが「本当かな？」と疑う姿を見て育った子どもは、自然に「本当かな？」と考えるようになる。言葉で教えるより、生き方で見せる方が100倍伝わる。それがお母さんという存在の、圧倒的な強さです。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例①：1年後の「ある家族の朝」", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "この思考法を実践したHさん（38歳）の、1年後の話をご紹介します。ある朝、小学6年生の息子がニュースを見ながらこう言いました。「お母さん、これって政府が言ってることだから、反対の意見も調べてみないとわからないよね」", spacing: { after: 200 } }),
        new Paragraph({ text: "1年前、息子はYouTubeで見た情報をそのまま話す子でした。それがいつの間にか「反対の意見も調べてみないと」と言えるようになった。Hさんは毎晩の夕食で「それって誰が言ってたの？」「あなたはどう思う？」を繰り返しただけです。「子どもに教えようとしていたのに、気づいたら私の方が教わっていました」Hさんはそう笑って話してくれました。", spacing: { after: 300 } }),

        new Paragraph({ text: "具体例②：変わったのは子どもだけじゃなかった", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "Iさん（42歳）はこの思考法を始めて3ヶ月後、夫との会話が変わっていたことに気づきました。以前は夫の意見に「そうだね」と合わせるだけだったIさんが、「でも私はこう思うんだけど、どう思う？」と言えるようになっていた。夫は「お前と話すの、最近面白いな」と言うようになったそうです。子どもに伝えようとして始めたことが、家族全体の空気を変えていた。「家族を守る思考法」は、お母さん自身も変える思考法でもあったんです。", spacing: { after: 300 } }),

        new Paragraph({ text: "今すぐ使える「最初の問いかけ」1フレーズ", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "明日の朝、子どもに言う言葉はこれだけで十分です。「それって、誰が言ってたの？」ニュースでも、友だちの話でも、YouTubeの内容でも。何か話してきたとき、この一言を添えるだけ。答えられなくてもいい。「わからない」でもいい。その「わからない」という気づきが、思考の始まりです。", spacing: { after: 300 } }),

        new Paragraph({ text: "この本を読んだあなたへ——家族を守る思考は連鎖する", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "あなたがこの思考法を身につけると、子どもに伝わります。子どもが身につけると、友だちに伝わります。友だちの家庭に広がると、地域が変わります。", spacing: { after: 200 } }),
        new Paragraph({ text: "AI時代に流されない家族をつくること。「本当に？」「なぜ？」「あなたはどう思う？」と問いかけられる親になること。それは子どものためだけじゃなく、あなた自身が「自分の頭で生きる」ための宣言でもあります。", spacing: { after: 300 } }),

        new Paragraph({ text: "おわりに｜「考える力」を守れるのは、お母さんだけです", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "明日の朝、子どもが起きてきたら。学校に行く前の、ほんの1分間。ニュースでも、昨日の出来事でも、なんでもいい。子どもが何か話したとき、この一言を言ってみてください。", spacing: { after: 200 } }),
        new Paragraph({
          children: [new TextRun({ text: "「それって、誰が言ってたの？」", bold: true, size: 28 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
        }),
        new Paragraph({ text: "その一言が、今日から家族の未来を変えていきます。難しくない。特別じゃない。あなたにできます。今日から、始められます。", spacing: { after: 200 } }),
        new Paragraph({ text: "最後まで読んでくれて、本当にありがとうございました。あなたの家族が、AI時代を自分の頭で、笑顔で生き抜いていけますように。", spacing: { after: 600 } }),

        // ── おわりに ──
        new Paragraph({ text: "おわりに｜あなたが変われば、家族が変わる。家族が変われば、未来が変わる。", heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } }),
        new Paragraph({ text: "最後まで読んでくれて、本当にありがとうございます。読み始めたとき、あなたはどんな気持ちでしたか？「AI時代についていけるか不安」「子どもに何も教えられないかもしれない」「自分には難しすぎるかな」そんな気持ちで開いてくれた方もいるかもしれません。でも今のあなたは、読む前と少し違うはずです。", spacing: { after: 300 } }),
        new Paragraph({ text: "この本でお伝えしたことは、たった5つの問いかけでした。", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "「それって、誰が言ってたの？」\n「本当にそうかな？他の見方はないかな？」\n「もし違ったら、どうなると思う？」\n「あなたはどう思う？」\n「それを知って、どうしたい？」", spacing: { after: 200 } }),
        new Paragraph({ text: "これだけです。特別な知識も、難しい技術も要らない。毎日完璧にやらなくていい。週に3回、思い出したときに1つ使うだけで十分です。でも、続けてください。思考習慣は、筋肉と同じです。使えば育ち、使わなければ落ちていく。そして一度育った思考の筋肉は、AI時代のどんな波が来ても、あなたと子どもを守り続けてくれます。", spacing: { after: 300 } }),
        new Paragraph({ text: "最後に、一つだけお願いがあります。", heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 200 } }),
        new Paragraph({ text: "この本の内容を、一人で抱えないでください。夫に話してみてください。「最近こんなことを読んでさ、うちもやってみようと思うんだけど」と。ママ友に話してみてください。「AIって便利だけど、使い方って大事だよね」と。あなたが話すことで、その言葉は誰かの心に届きます。誰かの家庭でも同じ問いかけが生まれます。考える力は、伝染します。良い意味で、広がっていきます。", spacing: { after: 300 } }),
        new Paragraph({ text: "AI時代を生き抜くのは、賢い人でも、ITに詳しい人でも、学歴がある人でもありません。「自分の頭で考えることをやめない人」です。そしてその姿を毎日子どもに見せられるのは、そばにいるお母さん、あなただけです。難しくない。特別じゃない。あなたにできます。", spacing: { after: 300 } }),
        new Paragraph({ text: "明日の朝、子どもが起きてきたら——ほんの一言、投げかけてみてください。", spacing: { after: 200 } }),
        new Paragraph({
          children: [new TextRun({ text: "「それって、誰が言ってたの？」", bold: true, size: 28 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
        }),
        new Paragraph({ text: "その一言が、あなたの家族の未来を変えていきます。応援しています。", spacing: { after: 400 } }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("C:/Users/owner/dev/note/AI時代の家族を守る思考法.docx", buffer);
console.log("完成しました！");
