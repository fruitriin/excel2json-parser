<template>
  <div id="app">
    <div class="file-drop-area" v-if="file">
      Excelファイルをドラッグ・アンド・ドロップしてください。
    </div>



  </div>
</template>

<script>
  import LandingPage from '@/components/LandingPage';

  const sheetConf = {
    headerRowOffset: 0,
      dataRowOffset: 0,
      
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
