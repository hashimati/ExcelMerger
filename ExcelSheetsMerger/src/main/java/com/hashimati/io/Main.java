package com.hashimati.io;

import com.hashimati.io.gui.ExcelMergerGUI;

import javax.swing.*;

public class Main {
    public static void main(String... args) throws ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {
        for(UIManager.LookAndFeelInfo info:UIManager.getInstalledLookAndFeels())
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    new ExcelMergerGUI();
                } catch (ClassNotFoundException e) {
                    e.printStackTrace();
                } catch (UnsupportedLookAndFeelException e) {
                    e.printStackTrace();
                } catch (InstantiationException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        });
    }
}
