<template>
  <div id="main">
    <div class="content">
      <div class="content-header">
        <div class="padding">
          <h1>Welcome</h1>
        </div>
      </div>
      <div class="content-main">
        <div class="padding">
          <button @click="onReload">Reload</button>
          <hr />
          <p>
            Choose the button below to set the color of the selected range to
            green.
          </p>
          <br />
          <button @click="onSetColor">Set color</button>
          <hr />
          <input
            ref="inputFile"
            type="file"
            @input="loadImage($event)"
            :style="styleObject"
          />
          <button @click="openFileDialog">Load image</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts">
import { defineComponent, ref } from "vue";

export default defineComponent({
  name: "App",
  data() {
    return {
      styleObject: { display: "none" },
    };
  },
  setup() {
    const inputFile = ref<HTMLInputElement | null>(null);

    return {
      inputFile
    }
  },
  methods: {
    onReload() {
      document.location.reload();
    },
    openFileDialog() {
      this.inputFile?.click();
    },
    loadImage($event: any) {
      const reader: any = new FileReader();

      reader.onload = () => {
        // @ts-ignore: Unreachable code error
        Excel.run(function (context: any) {
          const startIndex = reader.result.toString().indexOf("base64,");
          const myBase64 = reader.result.toString().substr(startIndex + 7);
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const image = sheet.shapes.addImage(myBase64);
          image.name = "Image";
          return context.sync();
        }).catch((err: unknown) => console.error(err));
      };

      // Read in the image file as a data URL.
      reader.readAsDataURL($event.target.files[0]);
    },
    onSetColor() {
      // @ts-ignore: Unreachable code error
      Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "green";
        await context.sync();
      });
    },
  },
});
</script>

<style>
.content-header {
  background: #2a8dd4;
  color: #fff;
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 80px;
  overflow: hidden;
}

.content-main {
  background: #fff;
  position: fixed;
  top: 80px;
  left: 0;
  right: 0;
  bottom: 0;
  overflow: auto;
}

.padding {
  padding: 15px;
}
</style>
