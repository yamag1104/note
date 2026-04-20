import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from "docx";
import { writeFileSync } from "fs";

const h1 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 600, after: 200 } });
const h2 = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 400, after: 150 } });
const body = (text, after = 200) => new Paragraph({ text, spacing: { after } });
const blank = () => new Paragraph({ text: "", spacing: { after: 100 } });

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "游明朝", size: 22 } },
    },
  },
  sections: [{
    children: [

      // タイトル
      new Paragraph({
        children: [new TextRun({ text: "note記事制作ワークフロー 完全版", bold: true, size: 36 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      }),
      new Paragraph({
        text: "リサーチ → コンセプト → 目次 → 本文 → テキスト化 → 有料設定 → SNS告知",
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),

      // 全体フロー
      h1("全体フロー"),
      body("STEP1  市場調査リサーチ"),
      body("STEP2  コンセプト整理"),
      body("STEP3  目次作成"),
      body("STEP4  本文執筆（はじめに・第1〜6章・おわりに）"),
      body("STEP5  Wordファイル出力"),
      body("STEP6  有料設定ライン・価格設定"),
      body("STEP7  SNS告知文作成・投稿", 400),

      // STEP1
      h1("STEP1｜市場調査リサーチ"),
      h2("確認する4つの条件（必ず順番に聞く）"),
      body("① 調査テーマ　　：AI副業／思考法 など"),
      body("② ターゲット層　：主婦／20代会社員／フリーランス など"),
      body("③ 販売形態　　　：eBook／テンプレート・ツール／オンライン講座"),
      body("④ 重点項目　　　：インサイト深掘り／競合分析重視／ブルーオーシャン発見", 300),
      h2("出力する4項目"),
      body("1. ターゲットの深い悩み（インサイト）× 5つ"),
      body("   └ 表層の悩みではなく「夜も眠れない本音の動機」"),
      body("2. 競合コンテンツの分析 × 3件"),
      body("   └ タイトル・価格帯・訴求ポイント・弱点"),
      body("3. 市場の空白（ブルーオーシャン）× 3つ"),
      body("   └ 競合がまだ触れていないニッチな需要"),
      body("4. 売れるパワーワード × 10個"),
      body("   └ ターゲットの感情を揺さぶる単語", 400),

      // STEP2
      h1("STEP2｜コンセプト整理"),
      h2("出力する4要素"),
      body("・タイトル案　　　　：5本（パワーワードを2つ以上含める）"),
      body("・解決する悩み　　　：インサイトから抽出した本音の問題"),
      body("・どうなれるか　　　：記事を読んだ後の読者の変化"),
      body("・独自の売り　　　　：競合にない切り口・市場の空白", 300),
      h2("タイトル作成の公式"),
      new Paragraph({
        children: [new TextRun({ text: "【ターゲット】＋【パワーワード①】＋【パワーワード②】", bold: true })],
        spacing: { after: 200 },
      }),
      body("例：「子どもに教えられるお母さんになる　AI時代の家族を守る思考法」", 400),

      // STEP3
      h1("STEP3｜目次作成"),
      h2("6章構成ルール"),
      body("第1章 問題提起　：読者の現状を言語化 →「私のことだ」と思わせる"),
      body("第2章 憧れ　　　：理想の姿を見せる →「こうなりたい」と思わせる"),
      body("第3章 解決策　　：具体的な方法を提示 →「これならできる」と思わせる"),
      body("第4章 証拠　　　：データ・実例で裏付け →「本当に効果がある」と信じさせる"),
      body("第5章 提案　　　：ステップ別アクション →「今日からできる」と感じさせる"),
      body("第6章 行動　　　：最初の一歩を促す →「やってみよう」と動かす", 300),
      body("※ 各章に「キャッチーな見出し」＋「2〜3個のサブ見出し」をつける", 400),

      // STEP4
      h1("STEP4｜本文執筆"),
      h2("執筆ルール"),
      body("・文字数    ：2,000文字程度／章"),
      body("・トーン    ：親しみやすく、少し熱量のある口調"),
      body("・構成      ：PREP法（結論→理由→具体例→結論）"),
      body("・具体例    ：失敗談・成功例を必ず1つずつ入れる"),
      body("・読者視点  ：「あなたは〇〇ではありませんか？」で引き込む", 300),
      h2("PREP法の構造"),
      body("【結論】  最初に「何を伝えたいか」をはっきり言う"),
      body("   ↓"),
      body("【理由】  「なぜそう言えるのか」根拠を説明する"),
      body("   ↓"),
      body("【具体例】 失敗談・成功例・データで実感させる"),
      body("   ↓"),
      body("【結論】  最初の結論を繰り返し、次章へ誘導する", 300),
      h2("執筆順序"),
      body("① はじめに（フック・本書の目的）"),
      body("② 第1章〜第6章（各2,000文字）"),
      body("③ おわりに（復習・継続の呼びかけ・最後の一押し）", 400),

      // STEP5
      h1("STEP5｜Wordファイル出力"),
      h2("実行コマンド"),
      new Paragraph({
        children: [new TextRun({ text: "cd C:\\Users\\owner\\dev\\note", font: "Courier New" })],
        spacing: { after: 100 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "node create_article.mjs", font: "Courier New" })],
        spacing: { after: 300 },
      }),
      h2("ファイル構成"),
      body("C:\\Users\\owner\\dev\\note\\"),
      body("├── create_article.mjs   ← Word生成スクリプト"),
      body("├── package.json"),
      body("├── node_modules\\"),
      body("└── {記事タイトル}.docx  ← 完成ファイル", 400),

      // STEP6
      h1("STEP6｜有料設定ライン・価格設定"),
      h2("有料ラインの引き方"),
      body("【無料】はじめに          ← 共感・フック"),
      body("【無料】第1章 問題提起    ← 「私のことだ」と思わせる"),
      body("【無料】第2章 憧れ（前半）← 「こうなりたい」を見せる"),
      new Paragraph({
        children: [new TextRun({ text: "━━━━━━ ここから有料 ━━━━━━", bold: true })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
      }),
      body("【有料】第2章 憧れ（後半）← 具体的な変化の姿"),
      body("【有料】第3章 解決策      ← 本書の核心（絶対有料）"),
      body("【有料】第4章〜第6章・おわりに", 300),
      h2("価格設定戦略"),
      body("初速（レビュー集め）：  980円　← 公開〜最初の1週間"),
      body("スタンダード　　　：1,980円　← レビュー5件集まったら"),
      body("実績あり　　　　　：2,980円　← SNS拡散・実績が出たら", 300),
      h2("有料ライン直前に入れる一文"),
      body("「ここから先では、〇〇が実際に使った"),
      body("『たった一言の問いかけ』と、"),
      body("△△が変わるまでの具体的なステップを公開しています。」", 400),

      // STEP7
      h1("STEP7｜SNS告知文"),
      h2("X投稿パターン別使い分け"),
      body("共感型　　　：公開初日　　→「私のことだ」で拡散"),
      body("問題提起型　：3日後　　　→ 危機感を煽って購買促進"),
      body("数字・実績型：レビュー後　→ 信頼性で背中を押す"),
      body("スレッド型　：土曜の朝　　→ 情報提供で拡散狙い", 300),
      h2("投稿タイミング"),
      body("X（平日）　　：21〜23時　← 子どもが寝た後にスマホを見る"),
      body("X（週末）　　：土曜 朝8〜9時　← 拡散されやすい時間帯"),
      body("Instagram　：20〜22時　← 週3回投稿が理想", 300),
      h2("公開後の動かし方"),
      body("DAY1   noteに記事を公開（980円）"),
      body("         → X・Instagramに共感型を投稿"),
      body("DAY3   X に問題提起型を投稿"),
      body("DAY7   X にスレッド型を投稿（土曜の朝）"),
      body("         レビュー5件集まり次第 → 1,980円に値上げ告知"),
      body("DAY14  値上げ告知をX・Instagramに投稿", 400),

      // チェックリスト
      h1("チェックリスト（次回記事用）"),
      h2("制作フェーズ"),
      body("□ STEP1：テーマ→ターゲット→販売形態→重点項目の順に確認してからリサーチ"),
      body("□ STEP2：パワーワードを2〜3個選んでタイトルに組み込む"),
      body("□ STEP3：タイトル確定後に目次を作成"),
      body("□ STEP4：はじめに→各章→おわりにの順で1章ずつ確認しながら執筆"),
      body("□ STEP5：create_article.mjsを新記事に書き換えてからnode実行", 300),
      h2("販売フェーズ"),
      body("□ STEP6：第2章の途中に有料ラインを設定／初速980円で公開"),
      body("□ STEP7：公開初日にX・Instagram投稿"),
      body("□ レビュー5件集まったら1,980円に値上げ＆告知投稿"),
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("C:/Users/owner/dev/note/note記事制作ワークフロー完全版.docx", buffer);
console.log("完成しました！");
