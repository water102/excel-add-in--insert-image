import { createApp } from 'vue'
import App from './App.vue'

// @ts-ignore: Unreachable code error
window.Office.onReady(() => {
    createApp(App).mount('#app');
});