<template>
  <div id="app">
    <div class="file-drop-area" v-if="!file">
      Excelファイルをドラッグ・アンド・ドロップしてください。
    </div>
    <div v-else>
    </div>

    <div class="sheetTabs">
      シート名(タブ)
    </div>
    <div class="sheetConfigs">
      シート別詳細設定

      出力列:

      ヘッダー行調整:
      データ行調整:

      データのキー:
      
    </div>


  </div>
</template>

<script>
  import LandingPage from '@/components/LandingPage';

  const sheetConfDefault = {
      willOutput: true,
      headerRowOffset: 0,
      dataRowOffset: 0,
      targetColNames: [],
      preview: false
  }

  export default {
      name: 'excel2json-parser',
      data: function(){
        return {
            file: {},
            rawData: {},
            config: {

            },
        }
      },
      computed:{
        previewData(){

        }
      },
      methods:{
          openFile(){

          },
          outputJson(){

          },
          saveConfig(){

          },
      }

  };
</script>

<style>
  /* CSS */
</style>
