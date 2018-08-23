<template>
  <div id="app">
    <div class="file-drop-area">
      Excelファイルをドラッグ・アンド・ドロップしてください。
    </div>

    <div class="sheetTabs">
      シート名(タブ)
    </div>
    <div class="sheetConfigs">

      シート別詳細設定<br>

      出力列:<br>

      ヘッダー行調整:<br>
      データ行調整:<br>

      データのキー:<br>

    </div>


  </div>
</template>

<script>
  const sheetConfDefault = {
    willOutput: true,
    headerRowOffset: 0,
    dataRowOffset: 0,
    targetColNames: [],
    preview: false,
  };

  export default {
    name: 'excel2json-parser',
    data() {
      return {
        file: {},
        rawData: {},
        config: {

        },
      };
  },
    computed: {
      previewData() {

      },
    },
    methods: {
      openFile() {

      },
      outputJson() {

      },
      saveConfig() {

      },
    },

  };
</script>

<style>
  /* CSS */
</style>
